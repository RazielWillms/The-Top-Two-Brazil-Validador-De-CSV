using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
                openFileDialog.InitialDirectory = @"C:\Users\Public";
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
            if (listBox1.SelectedIndex >= 0)
            {

                string filePath = txtFilePath.Text;

                if (File.Exists(filePath))
                {
                    try
                    {
                        DataTable dataTable = Importar_csv(filePath);
                        Validar_gerenciar(dataTable, listBox1.Text);
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

                if (possuiCabecalho)
                {
                    headers = primeiraLinha.Split(';');
                    //int colunas = headers.Length;
                    //Depura(colunas.ToString());
                }
                else
                {
                    headers = primeiraLinha.Split(';');
                    int colunas = headers.Length;
                    //Depura(colunas.ToString());
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

            //LC.Visible = true;
            //LC.Text = "C" + dataTable.Columns.Count.ToString() + "L" + dataTable.Rows.Count.ToString();

            return dataTable;
        }

        private void Validar_gerenciar(DataTable dataTable, String Tabela)
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

                default:
                    MessageBox.Show("A validação deste layout ainda não foi implementada", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        private void Exportar_click(object sender, EventArgs e)
        {
            Progresso_gerenciar(true);

            string filePath = @"C:\temp\RelatorioErros.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Erros");

                //insere o cabeçalho
                List<string[]> items = new List<string[]>
                {
                    new string[] { "Campo", "Linha", "Coluna", "Valor", "Obs" }
                };

                //obtém os erros da grid
                /*foreach (DataGridViewRow row in grid.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string[] values = new string[grid.Columns.Count];
                        for (int i = 0; i < grid.Columns.Count; i++)
                        {
                            values[i] = row.Cells[i].Value?.ToString();
                        }
                        items.Add(values);
                    }
                }*/

                //puxa da classe
                foreach (var registro in registros)
                {
                    items.Add(new string[] { registro.Campo, registro.Linha, registro.Coluna, registro.Valor, registro.Obs});
                }

                items = items.OrderBy(item => item[0])
                                .ThenBy(item => item[1])
                                .ThenBy(item => item[2])
                                .ToList();

                for (int i = 0; i < items.Count; i++)
                {
                    int total = items.Count;

                    for (int j = 0; j < items[i].Length; j++)
                    {
                        worksheet.Cell(i + 1, j + 1).Value = items[i][j];

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
            depuracao.Text = mensagem;
        }

        public void Grid_criar() 
        {
            grid.DataSource = null;
            grid.Rows.Clear();
            grid.Columns.Clear();

            labellog.Text = "Falhas encontradas: " + registros.Count;

            DataTable TableGrid = new DataTable();

            TableGrid.Rows.Clear();

            grid.AllowUserToOrderColumns = true;

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

        }
    }
}
