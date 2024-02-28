using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RenomeiaAgro
{
    public partial class TelaHome : Form
    {
        public string Propriedade { get; set; }

        public TelaHome()
        {
            InitializeComponent();
        }

        private void WHome_Load(object sender, EventArgs e)
        {
            lblUsuario.Text = CsUsuario.Nome;
        }

        private void BtnUpload_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            TelaUpload upload = new TelaUpload();
            upload.Show();
        }

        private void WHome_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }


    }
}
