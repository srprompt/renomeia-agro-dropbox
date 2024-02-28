using System.Collections.Generic;
using System.Windows.Forms;

namespace RenomeiaAgro
{
    public class CsArquivos
    {
        public static List<string> videos = new List<string>();
        public static List<string> documentos = new List<string>();

        public string AbrirSelecaoArquivoVideo()
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Selecionar CsArquivos para upload...",
                Multiselect = true,
                RestoreDirectory = true,
                Filter = "CsArquivos MP4 (*.mp4)|*.mp4",
                InitialDirectory = @"C:\Users\nomeUsuario"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            else
            {
                return null;
            }
        }

        public string AbrirSelecaoArquivoDocumento()
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Selecionar CsArquivos para upload...",
                Multiselect = true,
                RestoreDirectory = true,
                Filter = "CsArquivos PDF (*.pdf)|*.pdf",
                InitialDirectory = @"C:\Users\nomeUsuario"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            else
            {
                return null;
            }
        }
    }  
}
