using System;
using System.Windows.Forms;


namespace RenomeiaAgro
{
    public partial class TelaLogin : Form
    {

        public TelaLogin()
        {
            InitializeComponent();
        }

        private void BtnEntrar_Click(object sender, EventArgs e)
        {
            CsUsuario usuario = new CsUsuario();
            usuario.Acesso(txtLogin.Text, txtSenha.Text, this);
        }

        private void BtnLimpar_Click(object sender, EventArgs e)
        {
            txtLogin.Text = "";
            txtSenha.Text = "";
        }

        private void WLogin_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void TxtSenha_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CsUsuario usuario = new CsUsuario();
                usuario.Acesso(txtLogin.Text, txtSenha.Text, this);
            }
        }

    }
}
