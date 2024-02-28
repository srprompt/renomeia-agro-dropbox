using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RenomeiaAgro
{
    internal class CsExcel
    {
        public void CarregarPlanilha(ComboBox cbProfessor, ComboBox cbCurso, TextBox txtSigla)
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "BancoESM.xlsx");

            try
            {
                // Carrega o arquivo Excel dentro do bloco using
                using (var file = new System.IO.FileStream(dir, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(file);

                    // Seleciona a primeira planilha
                    ISheet sheet = workbook.GetSheetAt(0);

                    // Cria um HashSet para armazenar os valores únicos
                    HashSet<string> valoresUnicos = new HashSet<string>();

                    // Função para remover acentos, "ç" e pontos "." de uma string
                    string RemoverAcentosECaracteresEspeciais(string input)
                    {
                        string normalizedString = input.Normalize(NormalizationForm.FormD);
                        Regex regex = new Regex("[^a-zA-Z0-9 ]");
                        string withoutSpecialChars = regex.Replace(normalizedString, "");
                        return withoutSpecialChars;
                    }

                    // Percorre as linhas da coluna especificada e adiciona os valores ao ComboBox cbProfessor
                    for (int i = 1; i <= sheet.LastRowNum; i++) // Começa em 1 para ignorar a primeira linha
                    {
                        IRow row = sheet.GetRow(i);
                        if (row != null && row.GetCell(0) != null)
                        {
                            string valorCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(0).ToString());
                            cbProfessor.Items.Add(valorCelula);
                        }
                        if (row != null && row.GetCell(1) != null)
                        {
                            string valorCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(1).ToString());
                            valoresUnicos.Add(valorCelula);
                        }
                    }

                    foreach (string valor in valoresUnicos)
                    {
                        cbCurso.Items.Add(valor);
                    }

                    // Adicionar o evento de seleção de item para o ComboBox do curso
                    cbCurso.SelectedIndexChanged += (sender, e) =>
                    {
                        string cursoSelecionado = cbCurso.SelectedItem?.ToString();

                        // Limpar o ComboBox de professores
                        cbProfessor.Items.Clear();

                        // Percorrer as linhas da planilha para filtrar os professores do curso selecionado
                        for (int i = 1; i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row != null)
                            {
                                string cursoCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(1)?.ToString());
                                if (cursoCelula == cursoSelecionado)
                                {
                                    string professorCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(0)?.ToString());
                                    cbProfessor.Items.Add(professorCelula);
                                }
                            }
                        }

                        // Preencher o TextBox com a sigla do curso selecionado (celula do lado)
                        for (int i = 1; i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row != null)
                            {
                                string cursoCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(1)?.ToString());
                                if (cursoCelula == cursoSelecionado)
                                {
                                    string siglaCelula = RemoverAcentosECaracteresEspeciais(row.GetCell(2)?.ToString()); // Supondo que a sigla esteja na coluna 2 (terceira coluna)
                                    txtSigla.Text = siglaCelula;
                                    break; // Uma vez que encontramos a sigla, podemos parar de percorrer o resto da planilha
                                }
                            }
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }
        }
    }
}
