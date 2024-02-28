using System;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using static NPOI.HSSF.Record.UnicodeString;
using Dropbox.Api.Paper;
using Microsoft.SharePoint.Client.Discovery;
using OfficeOpenXml.Utils;
using static Dropbox.Api.Team.GroupAccessType;

namespace RenomeiaAgro
{
    public partial class TelaUpload : Form
    {
        int fimUpload = 0;
        private readonly string caminhoDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string Curso = "";

        public TelaUpload()
        {
            InitializeComponent();
        }

        private void WUpload_Load(object sender, EventArgs e)
        {
            lblUsuario.Text = CsUsuario.Nome;
            CsExcel excel = new CsExcel();
            excel.CarregarPlanilha(cbProfessor, cbCurso, txtSigla);
        }

        private void WUpload_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            var home = new TelaHome();
            home.FormClosed += (s, args) => this.Close();
            home.Visible = true;
        }

        private void MtxtDuracao_MouseClick(object sender, MouseEventArgs e)
        {
            mtxtDuracao.Focus();
            mtxtDuracao.SelectionStart = 0;
        }

        private string TransformaData() 
        {
            string dia = Convert.ToString(dtpData.Value.Day).PadLeft(2, '0');
            string mes = Convert.ToString(dtpData.Value.Month).PadLeft(2, '0');
            string ano = Convert.ToString(dtpData.Value.Year);
            string diaMesAno = dia + mes + ano;

            return diaMesAno;
        }


        private string TransformaSituacao()
        {
            if (cbSituacao.SelectedItem.ToString() == "Bruto (BR)")
            {
                return "BR";
            }
            else if (cbSituacao.SelectedItem.ToString() == "Editado (ED)")
            {
                return "ED";
            }
            return null;
        }

        private void CriarPastasVideos(string curso, string professor, string aulaData, string editadoBruto, string caminhoArq)
        {
            string raiz = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AGROADVANCE INOVACOES E TECNOLOGIA");
            string nomeProfessor = Path.Combine(raiz, RemoveInvalidFileNameChars(professor));
            string nomeAulaData = Path.Combine(nomeProfessor, aulaData);
            string situacao = Path.Combine(nomeAulaData, RemoveInvalidFileNameChars(editadoBruto));
            string documento = Path.Combine(nomeAulaData, "PDF");

            // Cria as pastas
            Directory.CreateDirectory(raiz);
            Directory.CreateDirectory(nomeProfessor);
            Directory.CreateDirectory(nomeAulaData);
            Directory.CreateDirectory(situacao);

            if (!Directory.Exists(documento))
            {
                Directory.CreateDirectory(documento);
            }

            // Move o arquivo para a última pasta criada
            string arqRenomVideo = caminhoArq; // Substitua pelo caminho do arquivo real
            string arquivoDestino = Path.Combine(situacao, Path.GetFileName(arqRenomVideo));
            File.Move(arqRenomVideo, arquivoDestino);
        }

        private void CriarPastasDocumentos(string curso, string professor, string aulaData, string editadoBruto, string caminhoArq)
        {
            string raiz = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AGROADVANCE INOVACOES E TECNOLOGIA");
            string nomeProfessor = Path.Combine(raiz, RemoveInvalidFileNameChars(professor));
            string nomeAulaData = Path.Combine(nomeProfessor, aulaData);
            string situacao = Path.Combine(nomeAulaData, RemoveInvalidFileNameChars(editadoBruto));
            string documento = Path.Combine(nomeAulaData, "PDF");

            // Cria as pastas
            Directory.CreateDirectory(raiz);
            Directory.CreateDirectory(nomeProfessor);
            Directory.CreateDirectory(nomeAulaData);
            Directory.CreateDirectory(situacao);
            Directory.CreateDirectory(documento);

            // Move o arquivo para a última pasta criada
            string arqRenomVideo = caminhoArq; // Substitua pelo caminho do arquivo real
            string arquivoDestino = Path.Combine(documento, Path.GetFileName(arqRenomVideo));
            File.Move(arqRenomVideo, arquivoDestino);
        }

        private string RemoveInvalidFileNameChars(string input)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            foreach (char c in invalidChars)
            {
                input = input.Replace(c.ToString(), "");
            }
            return input;
        }

        private void Renomear()
        {
            try
            {
                Curso = cbCurso.SelectedItem.ToString().ToUpper();
                string siglaCurso = txtSigla.Text.ToUpper();
                string data = TransformaData();
                string professor = cbProfessor.SelectedItem.ToString().ToUpper();
                string aula = txtAula.Text.ToUpper();
                string situacao = TransformaSituacao().ToUpper();
                string situacaoParaPasta = cbSituacao.SelectedItem.ToString().ToUpper();

                string[] novoNomeArray = new string[] { siglaCurso, data, professor, aula, "", situacao };
                string novoNomeString;

                for (int i = 0; i < CsArquivos.videos.Count; i++)
                {
                    string video = CsArquivos.videos[i].ToString();
                    string diretorio = System.IO.Path.GetDirectoryName(video);
                    string extensao = System.IO.Path.GetExtension(video);

                    //Altera valor do bloco conforme a quantidade de videos do Arraylist
                    int indice = i + 1;
                    novoNomeArray[4] = "BL" + indice;

                    //Transforma o Array para string
                    novoNomeString = string.Join("_", novoNomeArray);

                    string novoDiretorio = System.IO.Path.Combine(diretorio, novoNomeString + extensao);

                    System.IO.File.Move(video, novoDiretorio);

                    CsArquivos.videos[i] = novoDiretorio;

                    CriarPastasVideos(Curso, professor, aula + "_" + data, situacaoParaPasta, CsArquivos.videos[i].ToString());
                }

                for (int j = 0; j < CsArquivos.documentos.Count; j++)
                {
                    string documento = CsArquivos.documentos[j].ToString();
                    string diretorio = System.IO.Path.GetDirectoryName(documento);
                    string extensao = System.IO.Path.GetExtension(documento);

                    //Altera valor do bloco conforme a quantidade de videos do Arraylist
                    int indice = j + 1;
                    novoNomeArray[4] = "BL" + indice;

                    //Transforma o Array para string
                    novoNomeString = string.Join("_", novoNomeArray);

                    string novoDiretorio = System.IO.Path.Combine(diretorio, novoNomeString + extensao);

                    // Renomeie o arquivo
                    System.IO.File.Move(documento, novoDiretorio);

                    CsArquivos.documentos[j] = novoDiretorio;

                    CriarPastasDocumentos(Curso, professor, aula + "_" + data, situacaoParaPasta, CsArquivos.documentos[j].ToString());
                }

                //Console.WriteLine("Renomeação concluída.");
                MessageBox.Show("Pastas criadas e arquivos movidos com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                var reinicia = new CsReinicia();
                reinicia.LimpaArray();
                reinicia.LimpaCampos(this.Controls);
                reiniciaBtnVideos();
                reiniciaBtnDocumentos();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao renomear arquivo! " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }


        private void AtivaProximoBotao(System.Windows.Forms.Button proximoBtn)
        {
            proximoBtn.Enabled = true;
        }

        //-------------------VIDEO------------------------//
        CsArquivos arquivos = new CsArquivos();
        private void btnSel1_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco1.Text = dir;
            CsArquivos.videos.Add(txtBloco1.Text);
            if (txtBloco1.Text != "")
            {
                AtivaProximoBotao(btnSel2);
            }
        }

        private void btnSel2_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco2.Text = dir;
            CsArquivos.videos.Add(txtBloco2.Text);
            if (txtBloco2.Text != "")
            {
                AtivaProximoBotao(btnSel3);
            }
        }

        private void btnSel3_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco3.Text = dir;
            CsArquivos.videos.Add(txtBloco3.Text);
            if (txtBloco3.Text != "")
            {
                AtivaProximoBotao(btnSel4);
            }
        }

        private void btnSel4_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco4.Text = dir;
            CsArquivos.videos.Add(txtBloco4.Text);
            if (txtBloco4.Text != "")
            {
                AtivaProximoBotao(btnSel5);
            }
        }

        private void btnSel5_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco5.Text = dir;
            CsArquivos.videos.Add(txtBloco5.Text);
            if (txtBloco5.Text != "")
            {
                AtivaProximoBotao(btnSel6);
            }
        }

        private void btnSel6_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco6.Text = dir;
            CsArquivos.videos.Add(txtBloco6.Text);
            if (txtBloco6.Text != "")
            {
                AtivaProximoBotao(btnSel7);
            }
        }

        private void btnSel7_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco7.Text = dir;
            CsArquivos.videos.Add(txtBloco7.Text);
            if (txtBloco7.Text != "")
            {
                AtivaProximoBotao(btnSel8);
            }
        }

        private void btnSel8_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco8.Text = dir;
            CsArquivos.videos.Add(txtBloco8.Text);
            if (txtBloco8.Text != "")
            {
                AtivaProximoBotao(btnSel9);
            }
        }

        private void btnSel9_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco9.Text = dir;
            CsArquivos.videos.Add(txtBloco9.Text);
            if (txtBloco9.Text != "")
            {
                AtivaProximoBotao(btnSel10);
            }
        }

        private void btnSel10_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoVideo();
            txtBloco10.Text = dir;
            CsArquivos.videos.Add(txtBloco10.Text);
        }

        private void reiniciaBtnVideos()
        {
            btnSel1.Enabled = true;
            btnSel2.Enabled = false;
            btnSel3.Enabled = false;
            btnSel4.Enabled = false;
            btnSel5.Enabled = false;
            btnSel6.Enabled = false;
            btnSel7.Enabled = false;
            btnSel8.Enabled = false;
            btnSel9.Enabled = false;
            btnSel10.Enabled = false;
        }



        //------------------------------------------------//
        //-------------------DOCUMENTOS------------------//
        private void btnSel11_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco11.Text = dir;
            CsArquivos.documentos.Add(txtBloco11.Text);
            if (txtBloco11.Text != "")
            {
                AtivaProximoBotao(btnSel21);
            }
        }

        private void btnSel21_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco21.Text = dir;
            CsArquivos.documentos.Add(txtBloco21.Text);
            if (txtBloco21.Text != "")
            {
                AtivaProximoBotao(btnSel31);
            }
        }

        private void btnSel31_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco31.Text = dir;
            CsArquivos.documentos.Add(txtBloco31.Text);
            if (txtBloco31.Text != "")
            {
                AtivaProximoBotao(btnSel41);
            }
        }

        private void btnSel41_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco41.Text = dir;
            CsArquivos.documentos.Add(txtBloco41.Text);
            if (txtBloco41.Text != "")
            {
                AtivaProximoBotao(btnSel51);
            }
        }

        private void btnSel51_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco51.Text = dir;
            CsArquivos.documentos.Add(txtBloco51.Text);
            if (txtBloco51.Text != "")
            {
                AtivaProximoBotao(btnSel61);
            }
        }

        private void btnSel61_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco61.Text = dir;
            CsArquivos.documentos.Add(txtBloco61.Text);
            if (txtBloco61.Text != "")
            {
                AtivaProximoBotao(btnSel71);
            }
        }

        private void btnSel71_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco71.Text = dir;
            CsArquivos.documentos.Add(txtBloco71.Text);
            if (txtBloco71.Text != "")
            {
                AtivaProximoBotao(btnSel81);
            }
        }

        private void btnSel81_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco81.Text = dir;
            CsArquivos.documentos.Add(txtBloco81.Text);
            if (txtBloco81.Text != "")
            {
                AtivaProximoBotao(btnSel91);
            }
        }

        private void btnSel91_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco91.Text = dir;
            CsArquivos.documentos.Add(txtBloco91.Text);
            if (txtBloco91.Text != "")
            {
                AtivaProximoBotao(btnSel101);
            }
        }

        private void btnSel101_Click(object sender, EventArgs e)
        {
            string dir = arquivos.AbrirSelecaoArquivoDocumento();
            txtBloco101.Text = dir;
            CsArquivos.documentos.Add(txtBloco101.Text);
        }

        private void reiniciaBtnDocumentos()
        {
            btnSel11.Enabled = true;
            btnSel21.Enabled = false;
            btnSel31.Enabled = false;
            btnSel41.Enabled = false;
            btnSel51.Enabled = false;
            btnSel61.Enabled = false;
            btnSel71.Enabled = false;
            btnSel81.Enabled = false;
            btnSel91.Enabled = false;
            btnSel101.Enabled = false;
        }

        //------------------------------------------------//

        private async void btnProcessar_Click(object button, EventArgs e)
        {
            string pastaAlvo = "AGROADVANCE INOVACOES E TECNOLOGIA";
            string pastaFinal = "[ENVIADOS] Agroadvance Inovacoes e Tecnologia";

            string caminhoPasta = System.IO.Path.Combine(caminhoDesktop, pastaAlvo);
            string caminhoPastaFinal = System.IO.Path.Combine(caminhoDesktop, pastaFinal);

            // Verificar se um curso foi selecionado.
            if (cbCurso.SelectedItem == null)
            {
                MessageBox.Show("Selecione um curso.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar se um professor foi selecionado.
            if (cbProfessor.SelectedItem == null)
            {
                MessageBox.Show("Selecione um professor.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar se uma situação foi selecionada.
            if (cbSituacao.SelectedItem == null)
            {
                MessageBox.Show("Selecione uma situação.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar se o campo "Nome da Aula" não está vazio.
            if (string.IsNullOrWhiteSpace(txtAula.Text))
            {
                MessageBox.Show("O campo \"Nome da Aula\" não pode estar vazio.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar se o campo vídeo Bloco 1 não está vazio
            if (string.IsNullOrEmpty(txtBloco1.Text))
            {
                MessageBox.Show("Selecione ao menos um arquivo de vídeo para ser processado", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string caminhoDropbox = $"/" + cbCurso.SelectedItem.ToString().ToUpper();

            string CursoPLinkDrop = cbCurso.SelectedItem.ToString();
            string ProfPLinkDrop = cbProfessor.SelectedItem.ToString();

            Renomear();
            Console.WriteLine(Curso);
            btnProcessar.Enabled = false;

            var uploader = new CsDropboxUploader();

            //minimiza a janela do programa
            this.WindowState = FormWindowState.Minimized;

            //inicializa a autenticacao do Dropbox
            try
            {
                // Aguarda o resultado do método assíncrono
                await uploader.AuthorizeAndGetToken();

                btnProcessar.Visible = false;
                lblProgresso.Visible = true;
                pbBarraProgressoUp.Visible = true;

                //retorna a janela do programa
                this.WindowState = FormWindowState.Normal;
                this.Enabled = false;

                TelaUploadBox uploadBox = new TelaUploadBox();
                uploadBox.Show();

                Progress<(int, int)> progress = new Progress<(int, int)>();
                progress.ProgressChanged += (sender, data) =>
                {
                    //Calcula a porcentagem e atualiza o valor da barra de progresso
                    int progressPercentage = (int)((double)data.Item1 / data.Item2 * 100);
                    lblProgresso.Text = progressPercentage.ToString() + "%";
                    pbBarraProgressoUp.Value = progressPercentage;

                    //Chama o metodo do uploadBox para que os valores de progresso sejam atualizados na janela em questao
                    uploadBox.mostraPorcentagem(progressPercentage, $"{data.Item1}/{data.Item2}");
                };


                //Encaminha a pasta com os arquivos para o dropbox
                try
                {
                    int finalProgress = await uploader.UploadFolder(caminhoPasta, caminhoDropbox, progress);

                    fimUpload = 1;
                    uploadBox.verificaFimUpload(fimUpload);

                    //Cria pasta final e move os arquivos para ela
                    Directory.CreateDirectory(caminhoPastaFinal);
                    try
                    {
                        // Move a pasta gerada para a pasta de destino
                        MoverPastaComConteudo(caminhoPasta, caminhoPastaFinal, Curso);

                        // Exibe uma mensagem de sucesso
                        Console.WriteLine("Pasta movida com sucesso!");
                    }
                    catch (Exception ex)
                    {
                        // Exibe uma mensagem de erro em caso de falha
                        MessageBox.Show($"Erro ao mover a pasta: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    DesProgressoHabProcessar();
                    this.Enabled = true;
                    string pastaDropbox = "https://www.dropbox.com/work/Agroadvance Inovações e Tecnlogia/" + CursoPLinkDrop + "/" + ProfPLinkDrop;
                    Process.Start(pastaDropbox);
                }
                catch (Exception ex)
                {
                    // Exibe a mensagem de exceção em uma MessageBox
                    MessageBox.Show($"Ocorreu um erro ao fazer o upload: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }
            catch (Exception ex)
            {
                // Código para tratar a exceção
                MessageBox.Show($"Não foi possível autenticar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopiarDiretorioRecursivamente(DirectoryInfo origem, DirectoryInfo destino)
        {
            Directory.CreateDirectory(destino.FullName);

            foreach (FileInfo arquivo in origem.GetFiles())
            {
                arquivo.CopyTo(Path.Combine(destino.FullName, arquivo.Name), true);
            }

            foreach (DirectoryInfo subpasta in origem.GetDirectories())
            {
                DirectoryInfo novaSubpasta = destino.CreateSubdirectory(subpasta.Name);
                CopiarDiretorioRecursivamente(subpasta, novaSubpasta);
            }
        }

        private void MoverPastaComConteudo(string origem, string destino, string curso)
        {
            DirectoryInfo origemDir = new DirectoryInfo(origem);
            DirectoryInfo pastaCursoDir = new DirectoryInfo(Path.Combine(destino, curso));

            CopiarDiretorioRecursivamente(origemDir, pastaCursoDir);

            // Remover a pasta de origem após mover o conteúdo
            origemDir.Delete(true);
        }

        private void DesProgressoHabProcessar()
        {
            lblProgresso.Visible = false;
            pbBarraProgressoUp.Visible = false;
            btnProcessar.Visible = true;
            btnProcessar.Enabled = true;
        }

        private void txtAula_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Verifica se o caractere digitado é uma letra acentuada ou o caractere "ç".
            if (char.IsLetter(e.KeyChar) && (e.KeyChar > 127 || e.KeyChar < 32))
            {
                // Cancela a digitação do caractere.
                e.Handled = true;
            }
        }
    }
}
