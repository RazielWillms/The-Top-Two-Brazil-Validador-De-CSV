using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static iText.Kernel.Pdf.Colorspace.PdfSpecialCs;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using Ganss.Excel;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.ComponentModel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Emit;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.SS.Formula.Functions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace The_Top_Two_Brazil_Validador_De_CSV
{

    public partial class Main : Form
    {

        public Main()
        {
            InitializeComponent();
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
                        DataTable dataTable = Importa_csv(filePath);
                        Validar(dataTable, listBox1.Text);
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

        private void Exportar_click(object sender, EventArgs e)
        {
            progressBar.Value = 0;
            progressBar.Visible = true;

            string filePath = @"C:\temp\RelatorioMaquinas.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Maquinas");

                List<string[]> items = new List<string[]>();
                for (int i = 0; i < log.Items.Count; i++)
                {
                    string[] values = log.Items[i].ToString().Split(';');
                    items.Add(values);
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
                            Atualiza_progresso(total, i);
                        }
                    }
                }

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Salvar Relatório";
                    saveFileDialog.InitialDirectory = @"C:\temp";
                    saveFileDialog.FileName = "Relatorio" + listBox1.SelectedItem.ToString() + ".xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = saveFileDialog.FileName;
                    }
                    else
                    {
                        progressBar.Visible = false;
                        return;
                    }
                }

                workbook.SaveAs(filePath);
            }

            progressBar.Visible = false;
            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        private DataTable Importa_csv(string filePath)
        {
            DataTable dataTable = new DataTable();
            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] headers = null;
                if (possuiCabecalho.Checked)
                {
                    headers = sr.ReadLine().Split(';');
                }
                else
                {
                    headers = Enumerable.Range(1, 4).Select(i => "Coluna " + i).ToArray();
                }
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
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

        private void Validar(DataTable dataTable, String Tabela)
        {

            switch (Tabela)
            {
                case "Máquinas":
                    Maquinas(dataTable);
                    break;

                case "Saldos Máquinas":
                    Saldos_maquinas(dataTable);
                    break;

                case "Adiantamentos":
                    Adiantamentos(dataTable);
                    break;

                case "Orçamento Balcão":
                    Orcamento_balcao(dataTable);
                    break;

                case "Orçamento Oficina":
                    Orcamento_oficina(dataTable);
                    break;

                case "Estatísticas":
                    Estatisticas(dataTable);
                    break;

                case "Veículos Clientes":
                    Veiculos_clientes(dataTable);
                    break;

                case "Imobilizado Itens":
                    Imobilizado_itens(dataTable);
                    break;

                case "Imobilizado Saldos":
                    Imobilizado_saldos(dataTable);
                    break;

                case "Legado Financeiro":
                    Legado_financeiro(dataTable);
                    break;

                case "Legado Pagamentos":
                    Legado_pagamentos(dataTable);
                    break;

                case "Legado Pedidos":
                    Legado_pedidos(dataTable);
                    break;

                case "Legado Pedidos Itens":
                    Legado_pedidos_itens(dataTable);
                    break;

                case "Legado Movimentacao":
                    Legado_movimentacao(dataTable);
                    break;

                default:
                    MessageBox.Show("A validação deste layout ainda não foi implementada", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        static bool Validar_numeric(string valor, string tipo)
        {
            string pattern = "";

            switch (tipo)
            {

                case "5,2":
                    pattern = @"^\d{1,5}(,\d{1,2})?$";
                    break;

                case "6,2":
                    pattern = @"^\d{1,6}(,\d{1,2})?$";
                    break;

                case "8,1":
                    pattern = @"^\d{1,8}(,\d{1})?$";
                    break;

                case "8,2":
                    pattern = @"^\d{1,8}(,\d{1,2})?$";
                    break;

                case "10,1":
                    pattern = @"^\d{1,10}(,\d{1})?$";
                    break;

                case "12,2":
                    pattern = @"^\d{1,12}(,\d{1,2})?$";
                    break;

                case "12,4":
                    pattern = @"^\d{1,12}(,\d{1,4})?$";
                    break;

                case "15,2":
                    pattern = @"^\d{1,15}(,\d{1,2})?$";
                    break;

                case "15,3":
                    pattern = @"^\d{1,15}(,\d{1,3})?$";
                    break;

                case "15,4":
                    pattern = @"^\d{1,15}(,\d{1,4})?$";
                    break;

                case "16,2":
                    pattern = @"^\d{1,16}(,\d{1,2})?$";
                    break;

                case "16,3":
                    pattern = @"^\d{1,16}(,\d{1,3})?$";
                    break;

                case "16,4":
                    pattern = @"^\d{1,16}(,\d{1,4})?$";
                    break;

            }

            return Regex.IsMatch(valor, pattern);
        }

        private void Valida_campo(string tabela, string campo, int linha, int coluna, string tipo, double tamanho, Boolean obrigatorio)
        {
            if (obrigatorio)
            {
                if (campo.Trim() == "0" || campo.Trim() == "" || campo.Trim() == "null" || campo.Trim() == "NULL")
                {
                    Adicionar_registro(tabela, linha, coluna, campo, "Campo obrigatório");
                    return;
                }
            }

            switch (tipo)
            {
                case "char":
                    if (campo.Length > tamanho)
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Excede " + tamanho.ToString() + " caracter");
                    }
                    break;

                case "numeric":
                    if (!Validar_numeric(campo, tamanho.ToString().Replace('.', ',')))
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve estar no formato '" + tamanho.ToString().Replace('.', ',') + "'");
                    }
                    break;

                case "date":
                    if (!Validar_date(campo))
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve estar em um formato de data válido");
                    }
                    break;

                case "integer":
                    if (campo.Length > tamanho && !int.TryParse(campo, out _))
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser um número inteiro");
                    }
                    break;

                case "importado_nacional":
                    List<String> dom_importado_nacional = new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "", "null", "NULL" };
                    if (!dom_importado_nacional.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dom_importado_nacional);
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser estar dentre as opções: " + opcoes);
                    }
                    break;

                case "controla_estoque":
                    List<String> dom_controla_estoque = new List<String> { "S", "N", "", "null", "NULL" };
                    if (!dom_controla_estoque.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dom_controla_estoque);
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser estar dentre as opções: " + opcoes);
                    }
                    break;

                case "situacao":
                    List<String> dom_situacao = new List<String> { "A", "I", "", "null", "NULL" };
                    if (!dom_situacao.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dom_situacao);
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser estar dentre as opções: " + opcoes);
                    }
                    break;

                case "usado":
                    List<String> dom_usado = new List<String> { "1", "0", "", "null", "NULL" };
                    if (!dom_usado.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dom_usado);
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser estar dentre as opções: " + opcoes);
                    }
                    break;

                case "classe":
                    List<String> dom_classe = new List<String> { "N", "B", "", "null", "NULL" };
                    if (!dom_classe.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dom_classe);
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser estar dentre as opções: " + opcoes);
                    }
                    break;
            }
        }

        static bool Validar_date(string data)
        {
            string[] formatos = { "dd-MM-yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy" };
            DateTime date;
            return DateTime.TryParseExact(data, formatos, null, System.Globalization.DateTimeStyles.None, out date);
        }

        private void Adicionar_registro(string campo, int linha, int coluna, string valor, string obs)
        {
            log.Items.Add($"{campo};{(linha + 1)};{coluna};{valor};{obs}");
        }

        private void Atualiza_progresso(int total, int progresso)
        {
            int porcentagem = (progresso * 100) / total;
            progressBar.Value = porcentagem;
        }

        private void Maquinas(DataTable dataTable)
        {

            log.Items.Clear();
            labellog.Text = "Falhas encontradas (Máquinas):";
            log.Items.Add("Campo;Linha;Coluna;valor;observacao");

            progressBar.Value = 0;
            progressBar.Visible = true;
            int total = dataTable.Rows.Count;

            int rows = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código do Produto*
                            Valida_campo("Código do Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 2: //B - Descrição*
                            Valida_campo("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 3: //C - Descrição adicional do item*
                            Valida_campo("Descrição adicional do item", row[column].ToString(), rows, columns, "char", 1200, true);
                            break;

                        case 4: //D - Tipo de mercadoria(programa de excelência em gestão)
                            Valida_campo("Tipo de mercadoria", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 5: //E - Marca
                            Valida_campo("Marca", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 6: //F - Departamento
                            Valida_campo("Departamento", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 7: //G - Controla estoque
                            Valida_campo("Controla estoque", row[column].ToString(), rows, columns, "controla_estoque", 0, false);
                            break;

                        case 8: //H - Código do grupo*
                            Valida_campo("Departamento", row[column].ToString(), rows, columns, "integer", 10, true);
                            break;

                        case 9: //I - Peso liquido
                            Valida_campo("Pedo Liquido", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 10: //J - Peso bruto
                            Valida_campo("Peso bruto", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 11: //K - Unidade*
                            Valida_campo("Unidade", row[column].ToString(), rows, columns, "char", 2, true);
                            break;

                        case 12: //L - Aplicação
                            Valida_campo("Aplicação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 13: //M - Apelido
                            Valida_campo("Apelido", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 14: //N - Produto Importado ou Nacional
                            Valida_campo("Importado ou Nacional", row[column].ToString(), rows, columns, "importado_nacional", 0, false);
                            break;

                        case 15: //O - Preço de venda
                            Valida_campo("Preço de venda", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 16: //P - Preço de reposição
                            Valida_campo("Preço de reposição", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 17: //Q - Código de referência
                            Valida_campo("Código de referência", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 18: //R - Situação
                            Valida_campo("situacao", row[column].ToString(), rows, columns, "situacao", 0, false);
                            break;

                        case 19: //S - Produto usado*
                            Valida_campo("Produto usado", row[column].ToString(), rows, columns, "usado", 0, true);
                            break;

                        case 20: //T - NCM*
                            Valida_campo("NCM", row[column].ToString(), rows, columns, "char", 10, true);
                            break;

                        case 21: //U - Modelo
                            Valida_campo("Modelo", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 22: //V - Classe produto*
                            Valida_campo("Classe", row[column].ToString(), rows, columns, "classe", 0, true);
                            break;

                        case 23: //W - Código base*
                            Valida_campo("Código base", row[column].ToString(), rows, columns, "char", 20, true);
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Código base", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (row[column].ToString().Length > 20)
                                {
                                    Adicionar_registro("Código base", rows, columns, row[column].ToString(), "Excede 20 caracteres");
                                }
                            }
                            break;

                        case 24: //X - Número de serie
                            if (row[column].ToString().Length > 40)
                            {
                                Adicionar_registro("Número de serie", rows, columns, row[column].ToString(), "Excede 40 caracteres");
                            }
                            break;

                        case 25: //Y - Código antigo produto*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Código antigo", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (row[column].ToString().Length > 20)
                                {
                                    Adicionar_registro("Código antigo", rows, columns, row[column].ToString(), "Excede 20 caracteres");
                                }
                            }
                            break;

                        case 26: //Z - Código Fiscal
                            if (row[column].ToString().Length > 60)
                            {
                                Adicionar_registro("Código Fiscal", rows, columns, row[column].ToString(), "Excede 60 caracteres");
                            }
                            break;

                        case 27: //AB - Observação
                            if (row[column].ToString().Length > 1200)
                            {
                                Adicionar_registro("Observação", rows, columns, row[column].ToString(), "Excede 1200 caracteres");
                            }
                            break;

                        case 28: //AC - Controle de estoque*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Controle de estoque", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                List<String> dom_controle = new List<String> { "I" };
                                if (!dom_controle.Contains(row[column].ToString().Trim()))
                                {
                                    Adicionar_registro("Controle de estoque", rows, columns, row[column].ToString(), "Deve ser obrigatoriamente 'I'");
                                }
                            }
                            break;

                        case 29: //AD - Campo Livre
                            if (row[column].ToString().Length > 60)
                            {
                                Adicionar_registro("Campo Livre", rows, columns, row[column].ToString(), "Excede 60 caracteres");
                            }
                            break;

                        case 30: //AE - Filial*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro(column.ColumnName.ToString(), rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;

                        case 31: //AF - Código bandeira*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Código bandeira", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro(column.ColumnName.ToString(), rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;
                    }

                    if (columns > 31)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, " ", "Excedeu o número de colunas");
                    }

                    columns++;
                }
                Atualiza_progresso(total, rows);
                
                rows++;
            }

            progressBar.Visible = false;
        }

        private void Saldos_maquinas(DataTable dataTable)
        {
            log.Items.Clear();
            labellog.Text = "Falhas encontradas (Saldos Máquinas):";
            log.Items.Add("Campo;Linha;Coluna;valor;observacao");

            progressBar.Value = 0;
            progressBar.Visible = true;
            int total = dataTable.Rows.Count;

            int rows = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;

                        case 2: //B - Código do Produto*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Código do Produto", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (row[column].ToString().Length > 20)
                                {
                                    Adicionar_registro("Código do Produto", rows, columns, row[column].ToString(), "Excede 20 caracteres");
                                }
                            }
                            break;

                        case 3: //C - Quantidade*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Deve ser maior que zero");
                            }
                            else
                            {
                                if (!Validar_numeric(row[column].ToString(), "12,4"))
                                {
                                    Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Deve estar no formato '12,4', ex: 123,1234");
                                }
                            }
                            break;

                        case 4: //D - Valor do Estoque*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Valor do Estoque", rows, columns, row[column].ToString(), "Deve ser maior que zero");
                            }
                            else
                            {
                                if (!Validar_numeric(row[column].ToString(), "12,2"))
                                {
                                    Adicionar_registro("Valor do Estoque", rows, columns, row[column].ToString(), "Deve estar no formato '12,2', ex: 123,12");
                                }
                            }
                            break;

                        case 5: //E - Código da prateleira
                            if (row[column].ToString().Length > 10 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Código da prateleira", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 6: //F - Data da última compra
                            Valida_campo("Data da última compra", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 7: //G - Valor da última compra
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,4"))
                                {
                                    Adicionar_registro("Valor da última compra", rows, columns, row[column].ToString(), "Deve estar no formato '16,4', ex: 123,1234");
                                }
                            }
                            break;

                        case 8: //H - Estoque mínimo
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,4"))
                                {
                                    Adicionar_registro("Estoque mínimo", rows, columns, row[column].ToString(), "Deve estar no formato '16,4', ex: 123,1234");
                                }
                            }
                            break;

                        case 9: //I - Descrição
                            if (row[column].ToString().Length > 60)
                            {
                                Adicionar_registro("Descrição", rows, columns, row[column].ToString(), "Excede 60 caracteres");
                            }
                            break;

                        case 10: //J - Código produto único
                            if (row[column].ToString().Length > 10 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Código produto único", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 11: //K - Custo Reposição
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "15,2"))
                                {
                                    Adicionar_registro("Custo Reposição", rows, columns, row[column].ToString(), "Deve estar no formato '15,2', ex: 123,12");
                                }
                            }
                            break;

                        case 12: //L - Preço de venda
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "15,3"))
                                {
                                    Adicionar_registro("Preço de venda", rows, columns, row[column].ToString(), "Deve estar no formato '15,3', ex: 123,123");
                                }
                            }
                            break;
                    }

                    if (columns > 12)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, " ", "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualiza_progresso(total, rows);

                rows++;
            }

            progressBar.Visible = false;
        }
        
        private void Adiantamentos(DataTable dataTable)
        {
            log.Items.Clear();
            labellog.Text = "Falhas encontradas (Saldos Máquinas):";
            log.Items.Add("Campo;Linha;Coluna;valor;observacao");

            progressBar.Value = 0;
            progressBar.Visible = true;
            int total = dataTable.Rows.Count;

            int rows = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;

                        case 2: //B - Conta legado*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Conta legado", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (row[column].ToString().Length > 20)
                                {
                                    Adicionar_registro("Conta legado", rows, columns, row[column].ToString(), "Excede 20 caracteres");
                                }
                            }
                            break;

                        case 3: //C - Valor do adiantamento*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,2"))
                                {
                                    Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Deve estar no formato '16,2', ex: 123,12");
                                }
                            }
                            break;

                        case 4: //D - Tipo do adiantamento*

                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Tipo do adiantamento", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                List<String> dom_tipo_adiantamento = new List<String> { "C", "F" };
                                if (!dom_tipo_adiantamento.Contains(row[column].ToString().Trim()))
                                {
                                    Adicionar_registro("Tipo do adiantamento", rows, columns, row[column].ToString(), "Deve ser 'C' ou 'F'");
                                }
                            }
                            break;

                        case 5: //E - Centro de Custo
                            if (row[column].ToString().Length > 10 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Centro de Custo", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 6: //F - Número
                            if (row[column].ToString().Length > 10 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Número", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 7: //G - Observação
                            if (row[column].ToString().Length > 1200)
                            {
                                Adicionar_registro("Conta legado", rows, columns, row[column].ToString(), "Excede 1200 caracteres");
                            }
                            break;
                    }

                    if (columns > 7)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, " ", "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualiza_progresso(total, rows);

                rows++;
            }

            progressBar.Visible = false;
        }

        private void Orcamento_balcao(DataTable dataTable)
        {
            log.Items.Clear();
            labellog.Text = "Falhas encontradas (Saldos Máquinas):";
            log.Items.Add("Campo;Linha;Coluna;valor;observacao");

            progressBar.Value = 0;
            progressBar.Visible = true;
            int total = dataTable.Rows.Count;

            int rows = 1;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código Pedido*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro("Filial", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;

                        case 2: //B - Código do cliente (sistema antigo)*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Conta legado", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!decimal.TryParse(row[column].ToString(), out _))
                                {
                                    Adicionar_registro("Conta legado", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                                }
                            }
                            break;

                        case 3: //C - Operação
                            if (row[column].ToString().Length > 3 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Operação", rows, columns, row[column].ToString(), "Deve ser um número inteiro de no máximo 3 dígitos");
                            }
                            break;

                        case 4: //D - Politica Prazo
                            if (row[column].ToString().Length > 3 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Politica Prazo", rows, columns, row[column].ToString(), "Deve ser um número inteiro de no máximo 3 dígitos");
                            }
                            break;

                        case 5: //E - Politica Preço
                            if (row[column].ToString().Length > 3 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Politica Preço", rows, columns, row[column].ToString(), "Deve ser um número inteiro de no máximo 3 dígitos");
                            }
                            break;

                        case 6: //F - Tipo Operação
                            List<String> dom_tipo_operacao = new List<String> { "V", "S", "E", "C", "D" };
                            if (!dom_tipo_operacao.Contains(row[column].ToString().Trim()))
                            {
                                Adicionar_registro("Tipo Operação", rows, columns, row[column].ToString(), "Deve ser 'C' ou 'F'");
                            }
                            break;

                        case 7: //G - Vendedor
                            if (!decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Vendedor", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 8: //H - Funcionário Abertura O.C
                            if (!decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Funcionário Abertura O.C", rows, columns, row[column].ToString(), "Deve ser um número inteiro");
                            }
                            break;

                        case 9: //I - Data Validade
                            if (!Validar_date(row[column].ToString()) && row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                Adicionar_registro("Data Validade", rows, columns, row[column].ToString(), "Deve estar em um formato válido");
                            }
                            break;

                        case 10: //J - Data Abertura*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Data Abertura", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (!Validar_date(row[column].ToString()))
                                {
                                    Adicionar_registro("Data Abertura", rows, columns, row[column].ToString(), "Deve estar em um formato válido");
                                }
                            }
                            break;

                        case 11: //K - Data Parcelamento
                            if (!Validar_date(row[column].ToString()) && row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                Adicionar_registro("Data Parcelamento", rows, columns, row[column].ToString(), "Deve estar em um formato válido");
                            }
                            break;

                        case 12: //L - Situação*
                            List<String> dom_situacao = new List<String> { "A", "F" };
                            if (!dom_situacao.Contains(row[column].ToString().Trim()))
                            {
                                Adicionar_registro("Situação", rows, columns, row[column].ToString(), "Deve ser 'A' ou 'F'");
                            }
                            break;

                        case 13: //M - Status*
                            List<String> dom_status = new List<String> { "A", "P", "C", "F", "B", "S", "X", "Y" };
                            if (!dom_status.Contains(row[column].ToString().Trim()))
                            {
                                Adicionar_registro("Status", rows, columns, row[column].ToString(), "Deve ser 'A', 'P', 'C', 'F', 'B', 'S', 'X' ou 'Y'");
                            }
                            break;

                        case 14: //N - Produto*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Produto", rows, columns, row[column].ToString(), "Campo obrigatório");
                            }
                            else
                            {
                                if (row[column].ToString().Length > 20)
                                {
                                    Adicionar_registro("Produto", rows, columns, row[column].ToString(), "Excede 20 caracteres");
                                }
                            }
                            break;

                        case 15: //O - Descrição Produto
                            if (row[column].ToString().Length > 1200)
                            {
                                Adicionar_registro("Descrição Produto", rows, columns, row[column].ToString(), "Excede 1200 caracteres");
                            }
                            break;

                        case 16: //P - Quantidade*
                            if (row[column].ToString() == "0" || row[column].ToString().Trim() == "")
                            {
                                Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Deve ser maior que zero");
                            }
                            else
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,4"))
                                {
                                    Adicionar_registro("Quantidade", rows, columns, row[column].ToString(), "Deve estar no formato '16,4', ex: 123,1234");
                                }
                            }
                            break;

                        case 17: //Q - Preço Unitário*
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,3"))
                                {
                                    Adicionar_registro("Preço de venda", rows, columns, row[column].ToString(), "Deve estar no formato '16,3', ex: 123,123");
                                }
                            }
                            break;

                        case 18: //R - Valor Desconto
                            if (row[column].ToString() != "0" && row[column].ToString().Trim() != "")
                            {
                                if (!Validar_numeric(row[column].ToString(), "16,2"))
                                {
                                    Adicionar_registro("Preço de venda", rows, columns, row[column].ToString(), "Deve estar no formato '16,2', ex: 123,12");
                                }
                            }
                            break;

                        case 19: //S - Vendedor Produto
                            if (row[column].ToString().Length > 6 && !decimal.TryParse(row[column].ToString(), out _))
                            {
                                Adicionar_registro("Funcionário Abertura O.C", rows, columns, row[column].ToString(), "Deve ser um número inteiro de no máximo 6 dígitos");
                            }
                            break;
                    }

                    if (columns > 19)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, " ", "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualiza_progresso(total, rows);

                rows++;
            }

            progressBar.Visible = false;
        }

        private void Orcamento_oficina(DataTable dataTable)
        { 

        }

        private void Estatisticas(DataTable dataTable)
        { 
        }

        private void Veiculos_clientes(DataTable dataTable)
        { 
        }

        private void Imobilizado_itens(DataTable dataTable)
        {
        }

        private void Imobilizado_saldos(DataTable dataTable)
        {
        }

        private void Legado_financeiro(DataTable dataTable)
        {
        }

        private void Legado_pagamentos(DataTable dataTable)
        {
        }

        private void Legado_pedidos(DataTable dataTable)
        {
        }

        private void Legado_pedidos_itens(DataTable dataTable)
        {
        }
        private void Legado_movimentacao(DataTable dataTable)
        {
        }
    }
}
