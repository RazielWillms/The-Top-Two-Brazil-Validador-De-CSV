using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {

        private readonly List<Registro> registros;

        public Main()
        {
            InitializeComponent();
            registros = new List<Registro>();
            versao.Text = "v0.8";
        }

        public class Registro
        {
            public string Campo { get; set; }
            public string Linha { get; set; }
            public string Coluna { get; set; }
            public string Valor { get; set; }
            public string Obs { get; set; }

            public Registro(string campo, string linha, string coluna, string valor, string obs)
            {
                Campo = campo;
                Linha = linha;
                Coluna = coluna;
                Valor = valor;
                Obs = obs;
            }
        }

        public void Registro_adicionar(string campo, int linha, int coluna, string valor, string obs)
        {
            registros.Add(new Registro(campo, (linha + 1).ToString(), coluna.ToString(), valor, obs));
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

            if (layouts.SelectedIndex >= 0)
            {

                string filePath = txtFilePath.Text;

                if (File.Exists(filePath))
                {
                    try
                    {
                        DataTable dataTable = Importar_csv(filePath);
                        Validar_layouts_gerenciar(dataTable, layouts.Text);
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

        private DataTable Importar_csv(string filePath)
        {
            DataTable dataTable = new DataTable();
            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] headers = null;
                bool possuiCabecalho = this.possuiCabecalho.Checked;

                string primeiraLinha = sr.ReadLine() ?? throw new InvalidOperationException("O arquivo CSV está vazio.");

                //elimina campos inúteis ao final do arquivo
                string regex = "; {3,}"; //ponto e vírgula seguido de 3 ou mais espaços
                if (Regex.IsMatch(primeiraLinha, regex))
                {
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, ";");
                    regex = ";{3,}"; //3 ou mais ponto e vírgula seguidos
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, ";");
                    regex = " {2,}"; //2 ou mais espaços seguidos
                    primeiraLinha = Regex.Replace(primeiraLinha, regex, "");
                }

                if (possuiCabecalho)
                {
                    headers = primeiraLinha.Split(';');
                }
                else
                {
                    headers = primeiraLinha.Split(';');
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
                        primeiraLinhaDataRow[i] = primeiraLinhaDados[i];
                    }
                    dataTable.Rows.Add(primeiraLinhaDataRow);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(';');
                    DataRow dr = dataTable.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dataTable.Rows.Add(dr);
                }
            }

            return dataTable;
        }

        private void Validar_layouts_gerenciar(DataTable dataTable, String Tabela)
        {

            int rows = 0;

            if (possuiCabecalho.Checked)
            {
                rows = 1;
            }

            registros.Clear();

            switch (Tabela)
            {
                case "Máquinas":
                    Maquinas(dataTable, rows);
                    break;

                case "Saldos Máquinas":
                    Saldos_maquinas(dataTable, rows);
                    break;

                case "Adiantamentos":
                    Adiantamentos(dataTable, rows);
                    break;

                case "Orçamento Balcão":
                    Orcamento_balcao(dataTable, rows);
                    break;

                case "Orçamento Oficina":
                    Orcamento_oficina(dataTable, rows);
                    break;

                case "Estatísticas":
                    Estatisticas(dataTable, rows);
                    break;

                case "Veículos Clientes":
                    Veiculos_clientes(dataTable, rows);
                    break;

                case "Imobilizado Itens":
                    Imobilizado_itens(dataTable, rows);
                    break;

                case "Imobilizado Saldos":
                    Imobilizado_saldos(dataTable, rows);
                    break;

                case "Legado Financeiro":
                    Legado_financeiro(dataTable, rows);
                    break;

                case "Legado Pagamentos":
                    Legado_pagamentos(dataTable, rows);
                    break;

                case "Legado Pedidos":
                    Legado_pedidos(dataTable, rows);
                    break;

                case "Legado Pedidos Itens":
                    Legado_pedidos_itens(dataTable, rows);
                    break;

                case "Legado Movimentacao":
                    Legado_movimentacao(dataTable, rows);
                    break;

                case "Grupos":
                    if (NiveisCombo.Text == string.Empty)
                    {
                        MessageBox.Show("O campo Níveis da Empresa deve ser selecionado", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Grupos(dataTable, rows);
                    break;

                case "SubGrupos":
                    if (NiveisCombo.Text == string.Empty)
                    {
                        MessageBox.Show("O campo Níveis da Empresa deve ser selecionado", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (NivelCombo.Text == string.Empty)
                    {
                        MessageBox.Show("O campo Nível do Arquivo deve ser selecionado", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Sub_grupos(dataTable, rows);
                    break;

                default:
                    MessageBox.Show("A validação deste layout ainda não foi implementada", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
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

                Grid_criar();
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

            MensagemErro.Text = mensagem;
        }

        public void Grid_limpar()
        {
            grid.DataSource = null;
            grid.Rows.Clear();
            grid.Columns.Clear();

            labellog.Text = "Registro:";
        }

        public void Grid_criar()
        {
            Grid_limpar();

            if (registros.Count == 0)
            {
                labellog.Text = "Nenhum erro encontrado";
                MessageBox.Show("Nenhum erro encontrado", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                labellog.Text = "Erros: " + registros.Count;

                DataTable TableGrid = new DataTable();

                TableGrid.Rows.Clear();

                grid.AllowUserToOrderColumns = true;
                grid.ReadOnly = true;

                TableGrid.Columns.Add("Campo", typeof(string));
                TableGrid.Columns.Add("Linha", typeof(string));
                TableGrid.Columns.Add("Coluna", typeof(string));
                TableGrid.Columns.Add("Valor", typeof(string));
                TableGrid.Columns.Add("Observacao", typeof(string));

                foreach (var registro in registros)
                {
                    DataRow row = TableGrid.NewRow();
                    row["Campo"] = registro.Campo;
                    row["Linha"] = registro.Linha;
                    row["Coluna"] = registro.Coluna;
                    row["Valor"] = registro.Valor;
                    row["Observacao"] = registro.Obs;
                    TableGrid.Rows.Add(row);
                }

                grid.DataSource = TableGrid;

                Zoom_grid_criar();
            }
        }

        public void Zoom_grid_limpar()
        {
            btnZoomIn.Visible = false;
            btnZoomOut.Visible = false;
            zoom.Visible = false;
        }

        private void Zoom_grid_criar()
        {
            Zoom_grid_limpar();

            btnZoomIn.Visible = true;
            btnZoomOut.Visible = true;
            zoom.Visible = true;

            zoom.Text = "100%";
        }

        private void ZoomIn_click(object sender, EventArgs e)
        {
            Zoom_grid(grid, 2.0f);
        }

        private void ZoomOut_click(object sender, EventArgs e)
        {
            Zoom_grid(grid, -2.0f);
        }

        private void Zoom_grid(DataGridView dgv, float delta)
        {
            float currentFontSize = dgv.DefaultCellStyle.Font.Size;
            float newFontSize = currentFontSize + delta;

            if (newFontSize >= 6 && newFontSize <= 20)
            {
                dgv.DefaultCellStyle.Font = new System.Drawing.Font(dgv.DefaultCellStyle.Font.FontFamily, newFontSize);
                dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgv.ColumnHeadersDefaultCellStyle.Font.FontFamily, newFontSize);

                if (newFontSize > currentFontSize)
                {
                    Zoom_label_atualizar(20);
                }
                else
                {
                    Zoom_label_atualizar(-20);
                }
            }
        }

        private void Zoom_label_atualizar(int increment)
        {
            int currentZoom = int.Parse(zoom.Text.Replace('%', ' ').Trim());
            currentZoom += increment;
            zoom.Text = currentZoom.ToString() + '%';
        }

        private void Layout_selecionado(object sender, EventArgs e)
        {
            switch (layouts.Text)
            {
                case "Grupos":
                    Niveis.Visible = true;
                    NiveisCombo.Visible = true;
                    Nivel.Visible = false;
                    NivelCombo.Visible = false;
                    break;

                case "SubGrupos":
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
            NivelCombo.Items.Clear();
            NivelCombo.Items.Add("SubGrupo");
            NivelCombo.Items.Add("Segmento");
            NivelCombo.Items.Add("SubSegmento");

            if (layouts.Text == "SubGrupos")
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
