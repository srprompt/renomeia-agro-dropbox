using Dropbox.Api;
using Dropbox.Api.Files;
using Dropbox.Api.Auth;
using Dropbox.Api.Stone;
using Dropbox.Api.Users;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.Util.Collections;
using System.Net;
using System.Text;

namespace RenomeiaAgro
{
    public class CsDropboxUploader
    {
        private string AccessToken;
        private readonly string AppKey = Properties.Settings.Default.AppKey;
        private readonly string AppSecret = Properties.Settings.Default.AppSecret;
        private readonly string RedirectUri = Properties.Settings.Default.RedirectUri;

        public async Task<int> UploadFolder(string folderPath, string dropboxFolderPath, IProgress<(int, int)> progress)
        {
            int totalProgress = 0;

            try
            {
                using (var client = new DropboxClient(AccessToken))
                {
                    var folderInfo = new DirectoryInfo(folderPath);

                    long totalBytes = CalculateTotalBytes(folderInfo.GetFiles());
                    long uploadedBytes = 0;

                    foreach (var file in folderInfo.GetFiles())
                    {
                        if (file.Length > 150 * 1024 * 1024) // Check if the file is larger than 150MB
                        {
                            await ChunkUpload(client, dropboxFolderPath, file, progress);
                        }
                        else
                        {
                            using (var stream = file.OpenRead())
                            {
                                var dropboxPath = dropboxFolderPath + "/" + file.Name;

                                var uploadResult = await client.Files.UploadAsync(
                                    dropboxPath,
                                    WriteMode.Overwrite.Instance,
                                    body: stream
                                );

                                uploadedBytes += file.Length;
                                totalProgress = (int)(((double)uploadedBytes / totalBytes) * 100);

                                Console.WriteLine($"Progresso: {totalProgress}% concluído");
                                Console.WriteLine($"Arquivo enviado: {uploadResult.PathDisplay}");

                                progress.Report((totalProgress, 100)); // Since it's a single file, we report 100 as the total chunks.
                            }
                        }
                    }

                    foreach (var subfolder in folderInfo.GetDirectories())
                    {
                        var newDropboxFolderPath = dropboxFolderPath + "/" + subfolder.Name;
                        totalProgress = await UploadFolder(subfolder.FullName, newDropboxFolderPath, progress);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocorreu um erro: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return totalProgress;
        }

        private async Task ChunkUpload(DropboxClient client, string folder, FileInfo file, IProgress<(int, int)> progress)
        {
            Console.WriteLine("Chunk upload file...");

            // Chunk size é 125MB.
            const int chunkSize = 125 * 1024 * 1024;

            using (var stream = file.OpenRead())
            {
                int numChunks = (int)Math.Ceiling((double)stream.Length / chunkSize);

                byte[] buffer = new byte[chunkSize];
                string sessionId = null;

                int completedChunks = 0;

                for (var idx = 0; idx < numChunks; idx++)
                {
                    Console.WriteLine("Start uploading chunk {0}", idx);
                    var byteRead = stream.Read(buffer, 0, chunkSize);

                    using (MemoryStream memStream = new MemoryStream(buffer, 0, byteRead))
                    {
                        if (idx == 0)
                        {
                            var result = await client.Files.UploadSessionStartAsync(body: memStream);
                            sessionId = result.SessionId;
                        }
                        else
                        {
                            UploadSessionCursor cursor = new UploadSessionCursor(sessionId, (ulong)(chunkSize * idx));

                            if (idx == numChunks - 1)
                            {
                                try
                                {
                                    await client.Files.UploadSessionFinishAsync(cursor, new CommitInfo(folder + "/" + file.Name), body: memStream);
                                }
                                catch (Dropbox.Api.ApiException<Dropbox.Api.Files.UploadSessionFinishError> ex)
                                {
                                    // Verificar se a exceção é devido a espaço insuficiente.
                                    if (ex.Message.Contains("path/insufficient_space/"))
                                    {
                                        // Lógica para tratar o espaço insuficiente.
                                        // Por exemplo, exibir uma mensagem de erro ao usuário ou fazer outras ações apropriadas.
                                        MessageBox.Show("Espaço insuficiente no Dropbox. Não foi possível enviar o arquivo.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    else
                                    {
                                        // Se não for uma exceção de espaço insuficiente, exibir uma mensagem de erro com a descrição da exceção.
                                        MessageBox.Show($"Erro ao enviar o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        throw;
                                    }
                                }
                            }
                            else
                            {
                                await client.Files.UploadSessionAppendV2Async(cursor, body: memStream);
                            }
                        }
                    }

                    completedChunks++;

                    // Report the progress of completed chunks and total chunks.
                    progress.Report((completedChunks, numChunks));
                }
            }
        }

        private long CalculateTotalBytes(FileInfo[] files)
        {
            long totalBytes = 0;
            foreach (var file in files)
            {
                totalBytes += file.Length;
            }
            return totalBytes;
        }



        public async Task<string> AuthorizeAndGetToken()
        {
            var authorizeUri = DropboxOAuth2Helper.GetAuthorizeUri(OAuthResponseType.Code, AppKey, RedirectUri);

            // Inicia o processo de autorização
            Process.Start(authorizeUri.ToString());

            // Aguarda a resposta com o código de autorização
            var code = await WaitForAuthorizationCode();

            // Obtém o token de acesso
            AccessToken = await GetAccessToken(code);

            return AccessToken;
        }

        private async Task<string> WaitForAuthorizationCode()
        {
            var listener = new HttpListener();
            listener.Prefixes.Add(RedirectUri);

            try
            {
                listener.Start();
                Console.WriteLine("Aguardando código de autorização...");
                var context = await listener.GetContextAsync();
                var code = context.Request.QueryString.Get("code");
                Console.WriteLine($"Código de autorização recebido: {code}");

                // Retorna uma página HTML personalizada para o usuário
                var response = context.Response;
                response.ContentType = "text/html";
                var responseString = "<html><head><title>Autorização concluída</title></head><body><h1>Autorização concluída</h1><p>Autorização concluída. Você pode fechar esta página.</p></body></html>";
                var buffer = Encoding.UTF8.GetBytes(responseString);
                response.ContentLength64 = buffer.Length;
                var responseOutput = response.OutputStream;
                await responseOutput.WriteAsync(buffer, 0, buffer.Length);
                responseOutput.Close();
           
                return code;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocorreu um erro durante o processo de autorização: {ex.Message}");
                return null;
            }
            finally
            {
                listener.Stop();
            }
        }

        private async Task<string> GetAccessToken(string code)
        {
            var authResult = await DropboxOAuth2Helper.ProcessCodeFlowAsync(code, AppKey, AppSecret, RedirectUri);
            return authResult.AccessToken;
        }
    }
}
