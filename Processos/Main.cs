using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static ValidarCSV.TypeExtensions;

namespace ValidarCSV
{
    public partial class Main : Form
    {

        private readonly List<Registro> registros;

        public Main()
        {
            InitializeComponent();
            this.layouts.DataSource = new BindingSource(TypeExtensions.Layout_stringToEnum.Keys, null);
            registros = new List<Registro>();
            versao.Text = "v0.10";
        }

        private static string Numero_alfabeto_converter(int numero)
        {
            StringBuilder resultado = new StringBuilder();

            while (numero > 0)
            {
                numero--;
                resultado.Insert(0, (char)('A' + (numero % 26)));
                numero /= 26;
            }

            return resultado.ToString();
        }

        public void Registro_adicionar(string campo, int linha, int coluna, string valor, string obs)
        {
            registros.Add(new Registro(campo, (linha + 1).ToString(), Numero_alfabeto_converter(coluna), valor, obs));
        }

        private void Escolher_click(object sender, EventArgs e)
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (txtFilePath.Text == String.Empty)
                {
                    openFileDialog.InitialDirectory = @"C:\Users\Public";
                }
                else
                {
                    openFileDialog.InitialDirectory = txtFilePath.Text.ToString();
                }
                openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void Validar_click(object sender, EventArgs e)
        {
            Grid_limpar();

            Registro_gerenciar(true);
            Progresso_gerenciar(true);

            if (layouts.SelectedIndex >= 0)
            {

                string filePath = txtFilePath.Text;

                if (File.Exists(filePath))
                {
                    try
                    {
                        DataTable dataTable = Importar_csv(filePath);
                        Validar_layouts_gerenciar(dataTable, layouts.Text);

                        Progresso_gerenciar(false);
                        Registro_gerenciar(false);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao processar o arquivo: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Nenhum arquivo selecionado ou o arquivo não existe!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Layout Inválido!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable Importar_csv(string filePath)
        {
            DataTable dataTable = new DataTable();
            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] headers = null;
                bool possuiCabecalho = this.possuiCabecalho.Checked;

                string primeiraLinha = sr.ReadLine() ?? throw new InvalidOperationException("O arquivo CSV está vazio.");

                // Elimina campos inúteis ao final do arquivo
                string regex = "; {3,}"; // ponto e vírgula seguido de 3 ou mais espaços
                if (Regex.IsMatch(primeiraLinha, regex))
                {
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, ";");
                    regex = ";{3,}"; // 3 ou mais ponto e vírgula seguidos
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, ";");
                    regex = " {2,}"; // 2 ou mais espaços seguidos
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, "");
                }

                headers = primeiraLinha.Split(';');

                if (Repete_coluna(headers) || !possuiCabecalho)
                {
                    this.possuiCabecalho.Checked = false;

                    int colunas = headers.Length;
                    headers = Enumerable.Range(1, colunas).Select(i => "Coluna " + i).ToArray();
                }

                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                if (!possuiCabecalho)
                {
                    DataRow primeiraLinhaDataRow = dataTable.NewRow();
                    string[] primeiraLinhaDados = primeiraLinha.Split(';');
                    for (int i = 0; i < headers.Length; i++)
                    {
                        primeiraLinhaDataRow[i] = primeiraLinhaDados.Length > i ? primeiraLinhaDados[i] : "0";
                    }
                    dataTable.Rows.Add(primeiraLinhaDataRow);
                }

                int linha = 1;

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(';');
                    DataRow dr = dataTable.NewRow();

                    //verifica linhas menores que o esperado e as completa com valores nulos
                    if (rows.Length < headers.Length)
                    {
                        Registro_adicionar("Erro genérico", linha, (rows.Length + 1), "", $"Linha possui {rows.Length} colunas, menos que o esperado: {headers.Length}");

                        List<string> colunasCompletas = rows.ToList();
                        while (colunasCompletas.Count < headers.Length)
                        {
                            colunasCompletas.Add("0");
                        }
                        rows = colunasCompletas.ToArray();
                    }

                    //verifica as linhas com colunas sobressalentes e gera registro
                    if (rows.Length > headers.Length)
                    {
                        for (int i = (headers.Length - 1); i < rows.Length; i++)
                        {
                            Sobressalente_validar(linha, (i + 1), rows[i]);
                        }
                    }

                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dataTable.Rows.Add(dr);

                    linha++;
                }
            }
            return dataTable;
        }

        static bool Repete_coluna(string[] array)
        {
            if (array == null || array.Length == 0)
            {
                return false;
            }

            for (int i = 0; i < array.Length; i++)
            {
                for (int j = i + 1; j < array.Length; j++)
                {
                    if (array[i] == array[j])
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void Grid_datasource_alterado(object sender, EventArgs e)
        {
            //desabilita as ferramentas em torno da grid, exportar e zoom in e out
            if (grid.DataSource == null)
            {
                excel.Visible = false;
                Zoom_grid_limpar();

            }
            else
            {
                excel.Visible = true;
                Zoom_grid_criar();
            }
        }

        private void Exportar_click(object sender, EventArgs e)
        {
            Progresso_gerenciar(true);

            string filePath = @"C:\temp\RelatorioErros.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Erros");

                // insere o cabeçalho
                List<string[]> items = new List<string[]>
        {
            new string[] { "Campo", "Linha", "Coluna", "Valor", "Obs" }
        };

                // puxa da classe
                foreach (var registro in registros)
                {
                    items.Add(new string[] { registro.Campo, registro.Linha, registro.Coluna, registro.Valor, registro.Obs });
                }

                // ordena os itens sem o cabeçalho
                var sortedItems = items.Skip(1)
                                       .OrderBy(item => item[0])
                                       .ThenBy(item => item[1])
                                       .ThenBy(item => item[2])
                                       .ToList();

                // adiciona o cabeçalho novamente no início
                sortedItems.Insert(0, items[0]);

                for (int i = 0; i < sortedItems.Count; i++)
                {
                    int total = sortedItems.Count;

                    for (int j = 0; j < sortedItems[i].Length; j++)
                    {
                        worksheet.Cell(i + 1, j + 1).Value = sortedItems[i][j];

                        if (i % 250 == 0)
                        {
                            Progresso_atualizar(total, i);
                        }
                    }
                }

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Salvar Relatório";
                    saveFileDialog.InitialDirectory = @"C:\temp";
                    saveFileDialog.FileName = "Relatorio.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = saveFileDialog.FileName;
                    }
                    else
                    {
                        Progresso_gerenciar(false);
                        return;
                    }
                }

                workbook.SaveAs(filePath);
            }

            Progresso_gerenciar(false);
            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        public void Registro_gerenciar(bool Iniciar)
        {
            if (Iniciar)
            {
                registros.Clear();
            }
            else
            {
                Grid_criar();
            }
        }

        public void Progresso_gerenciar(bool Iniciar)
        {
            if (Iniciar)
            {
                progressBar.Value = 0;
                progressBar.Visible = true;
            }
            else
            {
                progressBar.Value = 0;
                progressBar.Visible = false;
            }
        }

        public void Progresso_atualizar(int total, int progresso)
        {
            int porcentagem = (progresso * 100) / total;
            progressBar.Value = porcentagem;
        }

        public void Mensagem_exibir(string mensagem)
        {
            depuracao.Visible = true;
            MensagemErro.Visible = true;

            MensagemErro.Text += mensagem;
        }

        private void Layout_selecionado(object sender, EventArgs e)
        {
            LayoutType layoutType = LayoutType.Indefinido;
            Layout_string_retornar(layouts.Text, ref layoutType);

            switch (layoutType)
            {
                case LayoutType.Grupos:
                    Niveis.Visible = true;
                    NiveisCombo.Visible = true;
                    Nivel.Visible = false;
                    NivelCombo.Visible = false;
                    break;

                case LayoutType.SubGrupos:
                    Niveis.Visible = true;
                    NiveisCombo.Visible = true;

                    if (NiveisCombo.Text != string.Empty)
                    {
                        Nivel.Visible = true;
                        NivelCombo.Visible = true;
                    }
                    break;

                default:
                    Nivel.Visible = false;
                    NivelCombo.Visible = false;
                    Niveis.Visible = false;
                    NiveisCombo.Visible = false;
                    break;
            }

        }

        private void NiveisCombo_selecionado(object sender, EventArgs e)
        {
            LayoutType layoutType = LayoutType.Indefinido;
            Layout_string_retornar(layouts.Text, ref layoutType);

            NivelCombo.Items.Clear();
            NivelCombo.Items.Add("SubGrupo");
            NivelCombo.Items.Add("Segmento");
            NivelCombo.Items.Add("SubSegmento");

            if (layoutType == LayoutType.SubGrupos)
            {
                Nivel.Visible = true;
                NivelCombo.Visible = true;
            }

            switch (NiveisCombo.Text)
            {
                case "2 (Grupo/SubGrupo)":
                    NivelCombo.Items.Remove("Segmento");
                    NivelCombo.Items.Remove("SubSegmento");
                    break;

                case "3 (Grupo/Subgrupo/Segmento)":
                    NivelCombo.Items.Remove("SubSegmento");
                    break;
            }

        }
    }
}
