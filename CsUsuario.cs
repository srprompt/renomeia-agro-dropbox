using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RenomeiaAgro
{
    public class CsUsuario
    {
        private readonly string LoginDB = Properties.Settings.Default.login;
        private readonly string SenhaDB = Properties.Settings.Default.senha;

        public static string Nome { get; set; }

        public void Acesso(string login, string senha, TelaLogin telaLogin)
        {
            if (login == LoginDB && senha == SenhaDB)
            {
                TelaHome home = new TelaHome();
                telaLogin.Hide();
                Nome = "Olá " + login.ToUpper() + "!";
                home.Show();
            }
            else
            {
                MessageBox.Show("Login ou senha inválidos", "Erro de Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
