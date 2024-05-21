using ClosedXML.Excel;
using MathNet.Numerics;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace The_Top_Two_Brazil_Validador_De_CSV
{

    public partial class Main : Form
    {
        private readonly List<Registro> registros;

        public Main()
        {
            InitializeComponent();
            registros = new List<Registro>();
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

            string filePath = @"C:\temp\RelatorioErros.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Erros");

                List<string[]> items = new List<string[]>();

                //obtém os erros no grid
                foreach (DataGridViewRow row in grid.Rows)
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
                }

                //habilitar caso queira puxar da classe
                // foreach (var registro in registros)
                // {
                //     items.Add(new string[] { registro.Nome, registro.Idade.ToString(), registro.Cidade });
                // }

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
                            Atualizar_progresso(total, i);
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
                        progressBar.Visible = false;
                        return;
                    }
                }

                workbook.SaveAs(filePath);
            }

            progressBar.Visible = false;
            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
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

            LC.Visible = true;
            LC.Text = "C" + dataTable.Columns.Count.ToString() + "L" + dataTable.Rows.Count.ToString();

            return dataTable;
        }

        private void Validar(DataTable dataTable, String Tabela)
        {

            int rows = 0;

            if (possuiCabecalho.Checked)
            {
                rows = 1;
            }

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

        public void Progresso(bool Iniciar)
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

        public void Adicionar_registro(string campo, int linha, int coluna, string valor, string obs)
        {
            registros.Add(new Registro($"{campo};", $"{(linha + 1)};", $"{coluna};",$"{valor};",$"{obs}"));
            //log.Items.Add($"{campo};{(linha + 1)};{coluna};{valor};{obs}");
        }

        public void Atualizar_progresso(int total, int progresso)
        {
            int porcentagem = (progresso * 100) / total;
            progressBar.Value = porcentagem;
        }

        public void Depura(string mensagem)
        {
            depuracao.Visible = true;
            depuracao.Text = mensagem;
        }

        public void Grid_criar() 
        {
            grid.DataSource = null;

            labellog.Text = "Falhas encontradas: " + registros.Count;

            DataTable TableGrid = new DataTable();

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

        //Validação
        public bool Valida_obrigatorio(string tabela, string campo, int linha, int coluna, string tipo)
        {

            if (campo.Trim() == "#" || campo.Trim() == "0" || campo.Trim() == "" || campo.Trim() == "null" || campo.Trim() == "NULL")
            {
                Adicionar_registro(tabela, linha, coluna, campo, "Campo obrigatório");
                return true;
            }

            if (string.IsNullOrEmpty(campo))
            {
                Adicionar_registro(tabela, linha, coluna, campo, "Campo está vazio");
                return true;
            }

            if (tipo == "integer")
            {
                if (Int32.TryParse(campo, out int valorInteiro))
                {
                    if (valorInteiro <= 0)
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser maior que zero");
                        return true;
                    }
                }
                else
                {
                    Adicionar_registro(tabela, linha, coluna, campo, "Formato inválido");
                    return true;
                }
            }

            if (tipo == "numeric")
            {
                if (decimal.TryParse(campo, out decimal valorDecimal))
                {
                    if (valorDecimal <= 0)
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve ser maior que zero");
                        return true;
                    }
                }
                else
                {
                    Adicionar_registro(tabela, linha, coluna, campo, "Formato inválido");
                    return true;
                }
            }

            return false;
        }

        public void Valida_dominio(string tabela, string campo, int linha, int coluna, List<String> dominio, Boolean obrigatorio)
        {
            if (obrigatorio)
            {
                if (Valida_obrigatorio(tabela, campo, linha, coluna, "N"))
                {
                    return;
                }
            }

            string opcoes;
            if (!dominio.Contains(campo.Trim()))
            {
                opcoes = String.Join(", ", dominio);
                Adicionar_registro(tabela, linha, coluna, campo, "Deve estar entre as opções: " + opcoes);
            }
        }

        public void Valida_campo(string tabela, string campo, int linha, int coluna, string tipo, double tamanho, Boolean obrigatorio)
        {
            if (obrigatorio)
            {
                if (Valida_obrigatorio(tabela, campo, linha, coluna, tipo))
                {
                    return;
                }
            }

            switch (tipo)
            {
                //campos padrão
                case "char":
                    if (campo.Length > tamanho)
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Excede " + tamanho.ToString() + " caracter");
                    }
                    break;

                case "numeric":
                    if (campo != "0" && campo.Trim() != "")
                    {
                        int parteInteira = (int)Math.Truncate(tamanho);
                        double parteDecimal = tamanho - parteInteira;
                        parteDecimal = parteDecimal.Round(1);
                        int parteFracionaria = (int)(parteDecimal * 10);

                        if (!Validar_numeric(campo.Trim(), parteInteira, parteFracionaria))
                        {
                            Adicionar_registro(tabela, linha, coluna, campo, "Deve estar no formato numérico: '" + tamanho.ToString().Replace('.', ',') + "'");
                        }
                    }
                    break;

                case "date":
                    if (!Validar_date(campo.Trim()))
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve estar em um formato de data válido");
                    }
                    break;

                case "date_format":
                    string formato = string.Empty;
                    Retorna_formato(tamanho, ref formato);

                    if (!Validar_date_format(campo.Trim(), formato))
                    {
                        Adicionar_registro(tabela, linha, coluna, campo, "Deve estar em um formato de data válido, conforme layout: " + formato);
                    }
                    break;

                case "integer":
                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho || !int.TryParse(campo, out _))
                        {
                            Adicionar_registro(tabela, linha, coluna, campo, "Deve ser um número inteiro e conter até " + tamanho + " dígitos");
                        }
                    }
                    break;

                default:
                    Adicionar_registro(tabela, linha, coluna, campo, "Validação falhou, conferir manualmente");
                    break;
            }
        }

        private bool Validar_numeric(string valor, int precisao, int escala)
        {
            if (valor == null || valor == "" || valor == "null" || valor == "NULL")
            {
                return true;
            }

            string pattern = @"^\d{1," + precisao.ToString().Trim() + @"}(,\d{1," + escala.ToString().Trim() + "})?$";
            return Regex.IsMatch(valor, pattern);
        }

        static bool Validar_date(string data)
        {
            if (data == null || data == "" || data == "null" || data == "NULL")
            { 
            return true; 
            }

            string[] formatos = { "dd-MM-yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy" };
            return DateTime.TryParseExact(data, formatos, null, System.Globalization.DateTimeStyles.None, out _);
        }

        private void Retorna_formato(double tipo, ref string formato) 
        {
            switch (tipo)
            {
                case 1:
                    formato = "dd-MM-yyyy";
                    break;

                case 2:
                    formato = "yyyy-MM-dd";
                    break;

                case 3:
                    formato = "yyyy/MM/dd";
                    break;

                case 4:
                    formato = "dd/MM/yyyy";
                    break;

                case 5:
                    formato = "yyyy-MM-dd HH:mm:ss";
                    break;

                case 6:
                    formato = "dd-MM-yyyy HH:mm:ss";
                    break;

                case 7:
                    formato = "yyyy/MM/dd HH:mm:ss";
                    break;

                case 8:
                    formato = "dd/MM/yyyy HH:mm:ss";
                    break;

                default:
                    formato = "NULL";
                    break;
            }
        }

        private bool Validar_date_format(string data, string formato)
        {
            if (data == null || data == "" || data == "null" || data == "NULL")
            {
                return true;
            }

            return DateTime.TryParseExact(data, formato, null, System.Globalization.DateTimeStyles.None, out _);
        }

        //layouts
        public void Maquinas(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

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
                            List<String> controla_estoque = new List<String> { "S", "N", "", "null", "NULL", "NULL" };
                            Valida_dominio("Controla estoque", row[column].ToString(), rows, columns, controla_estoque, false);
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
                            List<String> dom_importado_nacional = new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "", "null", "NULL", "NULL" };
                            Valida_dominio("Importado ou Nacional", row[column].ToString(), rows, columns, dom_importado_nacional, false);
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
                            List<String> dom_situacao = new List<String> { "A", "I", "", "null", "NULL", "NULL" };
                            Valida_dominio("situacao", row[column].ToString(), rows, columns, dom_situacao, false);
                            break;

                        case 19: //S - Produto usado*
                            List<String> dom_usado = new List<String> { "1", "0", "", "null", "NULL", "NULL" };
                            Valida_dominio("Produto usado", row[column].ToString(), rows, columns, dom_usado, true);
                            break;

                        case 20: //T - NCM*
                            Valida_campo("NCM", row[column].ToString(), rows, columns, "char", 10, true);
                            break;

                        case 21: //U - Modelo
                            Valida_campo("Modelo", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 22: //V - Classe produto*
                            List<String> dom_classe = new List<String> { "N", "B", "", "null", "NULL", "NULL" };
                            Valida_dominio("Classe", row[column].ToString(), rows, columns, dom_classe, true);
                            break;

                        case 23: //W - Código base*
                            Valida_campo("Código base", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 24: //X - Número de serie
                            Valida_campo("Número de serie", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 25: //Y - Código antigo produto*
                            Valida_campo("Código antigo produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 26: //Z - Código Fiscal
                            Valida_campo("Código Fiscal", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 27: //AB - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 28: //AC - Controle de estoque*
                            List<String> dom_controle = new List<String> { "I", "", "null", "NULL", "NULL" };
                            Valida_dominio("Controle de estoque", row[column].ToString(), rows, columns, dom_controle, true);
                            break;

                        case 29: //AD - Campo Livre
                            Valida_campo("Campo Livre", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 30: //AE - Filial*
                            Valida_campo("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 31: //AF - Código bandeira*
                            Valida_campo("Código bandeira", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;
                    }

                    if (columns > 31)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }
                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Saldos_maquinas(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Valida_campo("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Código do Produto*
                            Valida_campo("Código do Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 12.4, true);
                            break;

                        case 4: //D - Valor do Estoque*
                            Valida_campo("Valor do Estoque", row[column].ToString(), rows, columns, "numeric", 12.2, true);
                            break;

                        case 5: //E - Código da prateleira
                            Valida_campo("Código da prateleira", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 6: //F - Data da última compra
                            Valida_campo("Data da última compra", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 7: //G - Valor da última compra
                            Valida_campo("Valor da última compra", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 8: //H - Estoque mínimo
                            Valida_campo("Estoque mínimo", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 9: //I - Descrição
                            Valida_campo("Descrição", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 10: //J - Código produto único
                            Valida_campo("Código produto único", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 11: //K - Custo Reposição
                            Valida_campo("Estoque mínimo", row[column].ToString(), rows, columns, "numeric", 15.2, false);
                            break;

                        case 12: //L - Preço de venda
                            Valida_campo("Preço de venda", row[column].ToString(), rows, columns, "numeric", 15.3, false);
                            break;
                    }

                    if (columns > 12)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Adiantamentos(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Valida_campo("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Conta legado*
                            Valida_campo("Conta legado", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Campo Inutilizado
                            List<String> dom_inutilizado = new List<String> { "", "null", "NULL", "NULL" };
                            Valida_dominio("Campo Inutilizado", row[column].ToString(), rows, columns, dom_inutilizado, false);
                            break;

                        case 4: //D - Valor do adiantamento*
                            Valida_campo("Valor do adiantamento", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 5: //E - Tipo do adiantamento*
                            List<String> dom_tipo_adiantamento = new List<String> { "C", "F", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo do adiantamento", row[column].ToString(), rows, columns, dom_tipo_adiantamento, true);
                            break;

                        case 6: //F - Centro de Custo
                            Valida_campo("Centro de Custo", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: //G - Número
                            Valida_campo("Número", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 8: //H - Observação
                            Valida_campo("Conta legado", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;
                    }

                    if (columns > 8)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Orcamento_balcao(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código Pedido*
                            Valida_campo("Número", row[column].ToString(), rows, columns, "integer", 9, true);
                            break;

                        case 2: //B - Código do cliente (sistema antigo)*
                            Valida_campo("Código Legado do Cliente", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 3: //C - Operação
                            Valida_campo("Operação", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 4: //D - Política Prazo
                            Valida_campo("Política Prazo", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 5: //E - Politica Preço
                            Valida_campo("Politica Preço", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 6: //F - Tipo Operação
                            List<String> dom_tipo_operacao = new List<String> { "V", "S", "E", "C", "D", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo Operação", row[column].ToString(), rows, columns, dom_tipo_operacao, false);
                            break;

                        case 7: //G - Vendedor
                            Valida_campo("Vendedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 8: //H - Funcionário Abertura O.C
                            Valida_campo("Funcionário Abertura O.C", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 9: //I - Data Validade
                            Valida_campo("Data Validade", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 10: //J - Data Abertura*
                            Valida_campo("Data Abertura", row[column].ToString(), rows, columns, "date", 0, true);
                            break;

                        case 11: //K - Data Parcelamento
                            Valida_campo("Data Parcelamento", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 12: //L - Situação*
                            List<String> dom_orc_situacao = new List<String> { "A", "F", "", "null", "NULL", "NULL" };
                            Valida_dominio("Situação", row[column].ToString(), rows, columns, dom_orc_situacao, true);
                            break;

                        case 13: //M - Status*
                            List<String> dom_status = new List<String> { "A", "P", "C", "F", "B", "S", "X", "Y", "", "null", "NULL", "NULL" };
                            Valida_dominio("Status", row[column].ToString(), rows, columns, dom_status, true);
                            break;

                        case 14: //N - Produto*
                            Valida_campo("Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 15: //O - Descrição Produto
                            Valida_campo("Descrição Produto", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 16: //P - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 17: //Q - Preço Unitário*
                            Valida_campo("Preço Unitário", row[column].ToString(), rows, columns, "numeric", 16.3, true);
                            break;

                        case 18: //R - Valor Desconto
                            Valida_campo("Valor Desconto", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 19: //S - Vendedor Produto
                            Valida_campo("Vendedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;
                    }

                    if (columns > 19)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Orcamento_oficina(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Número*
                            Valida_campo("Número", row[column].ToString(), rows, columns, "integer", 9, true);
                            break;

                        case 2: //B - Código da Filial Solution*
                            Valida_campo("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 3: //C - ID do Veículo*
                            Valida_campo("ID do Veículo", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 4: //D - Série do veículo*
                            Valida_campo("Série do veículo", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 5: //E - Conta do cliente legado - sistema antigo*
                            Valida_campo("Conta do cliente legado", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 6: //F - Tipo da OS
                            Valida_campo("Tipo da OS", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: //G - Data de abertura
                            Valida_campo("Data de abertura", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 8: //H - ID do mecânico no Solution
                            Valida_campo("Mecânico no Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 9: //I - ID do vendedor no Solution
                            Valida_campo("Vendedor no Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 10: //J - ID do local de venda
                            Valida_campo("local de venda", row[column].ToString(), rows, columns, "integer", 6, true);
                            break;

                        case 11: //K - ID da política de preço
                            Valida_campo("política de preço", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 12: //L - ID da política de prazo
                            Valida_campo("política de prazo", row[column].ToString(), rows, columns, "char", 3, true);
                            break;

                        case 13: //M - Código do produto*
                            Valida_campo("Código do produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 14: //N - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 15: //O - Preço unitário*
                            Valida_campo("Preço unitário", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;
                    }

                    if (columns > 15)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Estatisticas(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código filial Solution*
                            Valida_campo("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Código produto*
                            Valida_campo("Código produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Data movimetação (mês e ano)*
                            Valida_campo("Data movimetação", row[column].ToString(), rows, columns, "date", 0, true);
                            break;

                        case 4: //D - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 15.4, true);
                            break;

                        case 5: //E - Valor total*
                            Valida_campo("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;
                    }

                    if (columns > 5)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Veiculos_clientes(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código*
                            Valida_campo("Código", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 2: // B - Descrição*
                            Valida_campo("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 3: // C - Placa
                            Valida_campo("Placa", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 4: // D - Meses Garantia
                            Valida_campo("Meses Garantia", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 5: // E - Hrs.Garantia
                            Valida_campo("Hrs.Garantia", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 6: // F - Km garantia
                            Valida_campo("Km garantia", row[column].ToString(), rows, columns, "numeric", 10.1, false);
                            break;

                        case 7: // G - Novo Usado*
                            List<String> dom_novo_usado = new List<String> { "N", "U", "", "null", "NULL", "NULL" };
                            Valida_dominio("Novo Usado", row[column].ToString(), rows, columns, dom_novo_usado, true);
                            break;

                        case 8: // H - Versão
                            Valida_campo("Versão", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 9: // I - Ano fabricação*
                            Valida_campo("Ano fabricação", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 10: // J - Ano modelo*
                            Valida_campo("Ano modelo", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 11: // K - Código da conta de cliente (sistema antigo)*
                            Valida_campo("Código da conta de cliente (sistema antigo)", row[column].ToString(), rows, columns, "char", 6, true);
                            break;

                        case 12: // L - Modelo*
                            Valida_campo("Modelo", row[column].ToString(), rows, columns, "char", 12, true);
                            break;

                        case 13: // M - numero NF de compra
                            Valida_campo("numero NF de compra", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 14: // N - Data de compra
                            Valida_campo("Data de compra", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 15: // O - Código da conta de fornecedor
                            Valida_campo("Código da conta de fornecedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 16: // P - Nome fornecedor
                            Valida_campo("Nome fornecedor", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 17: // Q - Código produto estoque
                            Valida_campo("Código produto estoque", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 18: // R - Numero de serie*
                            Valida_campo("Numero de serie", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 19: // S - Serie motor*
                            Valida_campo("Serie motor", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 20: // T - Série da bomba hidráulica
                            Valida_campo("Série da bomba hidráulica", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 21: // U - Série de transmissão
                            Valida_campo("Série de transmissão", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 22: // V - Série da caixa de câmbio
                            Valida_campo("Série da caixa de câmbio", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 23: // W - Série da bomba injetora
                            Valida_campo("Série da bomba injetora", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 24: // X - Série do monobloco
                            Valida_campo("Série do monobloco", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Série do eixo dianteiro
                            Valida_campo("Série do eixo dianteiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 26: // Z - Série da plataforma
                            Valida_campo("Série da plataforma", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 27: // AA - Pneus dianteiro
                            Valida_campo("Pneus dianteiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 28: // AB - Pneus traseiro
                            Valida_campo("Pneus traseiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 29: // AC - Série direção hidráulica
                            Valida_campo("Série direção hidráulica", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 30: // AD - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 200, false);
                            break;

                        case 31: // AE - Tipo equipamento*
                            List<String> dom_tipo_equipamento = new List<String> { "#", "J", "8", "4", "A", "5", "N", "C", "R", "D", "2", "L", "K", "P", "H", "V", "I", "3", "S", "6", "M", "O", "9", "Z", "B", "U", "F", "7", "Y", "T", "G", "Q", "1", "E", "X", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo equipamento", row[column].ToString(), rows, columns, dom_tipo_equipamento, true);
                            break;

                        case 32: // AF - Código do pedido da gestão de compra
                            Valida_campo("Código do pedido da gestão de compra", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 33: // AG - Cor código*
                            Valida_campo("Cor código", row[column].ToString(), rows, columns, "char", 4, true);
                            break;

                        case 34: // AH - Cor descrição*
                            Valida_campo("Cor descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 35: // AI - Potência do Motor (CV)
                            Valida_campo("Potência do Motor (CV)", row[column].ToString(), rows, columns, "numeric", 8.1, false);
                            break;

                        case 36: // AJ - CM3 (cilindradas)
                            Valida_campo("CM3 (cilindradas)", row[column].ToString(), rows, columns, "numeric", 8.1, false);
                            break;

                        case 37: // AK - Peso líquido (KG)
                            Valida_campo("Peso líquido (KG)", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 38: // AL - Peso bruto (KG)
                            Valida_campo("Peso bruto (KG)", row[column].ToString(), rows, columns, "numeric", 10, false);
                            break;

                        case 39: // AM - Tipo combustivel*
                            Valida_campo("Tipo combustivel", row[column].ToString(), rows, columns, "char", 10, true);
                            break;

                        case 40: // AN - CMKG
                            Valida_campo("CMKG", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 41: // AO - TMA
                            Valida_campo("TMA", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 42: // AP - Distância entre eixos (mm)
                            Valida_campo("Distância entre eixos (mm)", row[column].ToString(), rows, columns, "numeric", 8.2, false);
                            break;

                        case 43: // AQ - RENAVAM
                            Valida_campo("RENAVAM", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 44: // AR - Tipo pintura*
                            Valida_campo("Tipo pintura", row[column].ToString(), rows, columns, "char", 1, true);
                            break;

                        case 45: // AS - Tipo de Veículo Renavam/Denatran
                            List<String> dom_tipo_renavam_denatram = new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, dom_tipo_renavam_denatram, false);
                            break;

                        case 46: // AT - Espécie de Veículo Renavam/Denatran
                            List<String> dom_especie_veiculo_renavam_denatram = new List<String> { "0", "1", "2", "3", "4", "5", "6", "", "null", "NULL", "NULL" };
                            Valida_dominio("Espécie de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, dom_especie_veiculo_renavam_denatram, false);
                            break;

                        case 47: // AU - Marca Modelo Renavam/Denatran
                            Valida_campo("Marca Modelo Renavam/Denatran", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 48: // AV - Codigo do DN
                            Valida_campo("Codigo do DN", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 49: // AW - Chassis*
                            Valida_campo("Chassis", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 50: // AX - Marca
                            Valida_campo("Marca", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 51: // AY - Data entrega tecnica
                            Valida_campo("Data entrega tecnica", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 52: // AZ - Data ultima revisão
                            Valida_campo("Data ultima revisão", row[column].ToString(), rows, columns, "date", 0, false);
                            break;
                    }

                    if (columns > 52)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Imobilizado_itens(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código da Empresa Solution*
                            Valida_campo("Código da Empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código da Filial Solution*
                            Valida_campo("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 3: // C - Código do Item*
                            Valida_campo("Código do Item", row[column].ToString(), rows, columns, "numeric", 6.2, true);
                            break;

                        case 4: // D - Código da Conta (Plano de Contas)
                            Valida_campo("Código da Conta (Plano de Contas)", row[column].ToString(), rows, columns, "char", 11, false);
                            break;

                        case 5: // E - Data do lancto*
                            Valida_campo("Data do lancto", row[column].ToString(), rows, columns, "date", 10, true);
                            break;

                        case 6: // F - Data da aquisição*
                            Valida_campo("Data da aquisição", row[column].ToString(), rows, columns, "date", 10, true);
                            break;

                        case 7: // G - Centro de Custo
                            Valida_campo("Centro de Custo", row[column].ToString(), rows, columns, "char", 6, false);
                            break;

                        case 8: // H - % de Depreciação do Item
                            Valida_campo("% de Depreciação do Item", row[column].ToString(), rows, columns, "numeric", 5.2, false);
                            break;

                        case 9: // I - % de Depreciação Gerencial
                            Valida_campo("% de Depreciação Gerencial", row[column].ToString(), rows, columns, "numeric", 6.2, false);
                            break;

                        case 10: // J - % residual
                            Valida_campo("% residual", row[column].ToString(), rows, columns, "numeric", 5.2, false);
                            break;

                        case 11: // K - Débito ou Crédito*
                            List<String> dom_debito_credito = new List<String> { "D", "C", "", "null", "NULL", "NULL" };
                            Valida_dominio("Débito ou Crédito", row[column].ToString(), rows, columns, dom_debito_credito, true);
                            break;

                        case 12: // L - Chave*
                            List<String> dom_chave = new List<String> { "G", "C", "", "null", "NULL", "NULL" };
                            Valida_dominio("Chave", row[column].ToString(), rows, columns, dom_chave, true);
                            break;

                        case 13: // M - Tipo lançamento
                            List<String> dom_tipo_lanacamento = new List<String> { "A", "T", "I", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo lançamento", row[column].ToString(), rows, columns, dom_tipo_lanacamento, false);
                            break;

                        case 14: // N - Tipo Baixa
                            List<String> dom_tipo_baixa = new List<String> { "B", "T", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo Baixa", row[column].ToString(), rows, columns, dom_tipo_baixa, false);
                            break;

                        case 15: // O - Número do documento de aquisição
                            Valida_campo("Número do documento de aquisição", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 16: // P - Nome do Fornecedor
                            Valida_campo("Nome do Fornecedor", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 17: // Q - Descrição*
                            Valida_campo("Descrição", row[column].ToString(), rows, columns, "char", 225, true);
                            break;

                        case 18: // R - Descrição sucienta da função do bem na atividade do estabelecimento (obrigatório para Sped Fiscal)*
                            Valida_campo("Descrição sucienta", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 19: // S - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 20: // T - Número da Apólice
                            Valida_campo("Número da Apólice", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 21: // U - Data do Vencimento
                            Valida_campo("Data do Vencimento", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 22: // V - Código Externo
                            Valida_campo("Código Externo", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 23: // W - Código do Local
                            Valida_campo("Código do Local", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 24: // X - Código do Responsável
                            Valida_campo("Código do Responsável", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 25: // Y - Código do tipo do bem
                            Valida_campo("Código do tipo do bem", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 26: // Z - Código da Seguradora
                            Valida_campo("Código da Seguradora", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 27: // AA - Tipo Documento de aquisição
                            Valida_campo("Tipo Documento de aquisição", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 28: // AB - Situação do Bem
                            Valida_campo("Situação do Bem", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 29: // AC - Chassis
                            Valida_campo("Chassis", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 30: // AD - Placa
                            Valida_campo("Placa", row[column].ToString(), rows, columns, "char", 9, false);
                            break;
                    }

                    if (columns > 30)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Imobilizado_saldos(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código da Empresa*
                            Valida_campo("Código da Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código do Item*
                            Valida_campo("Código do Item", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 3: // C - Valor Original*
                            Valida_campo("Valor Original", row[column].ToString(), rows, columns, "numeric", 15.2, true);
                            break;

                        case 4: // D - Valor Original Corrigido
                            Valida_campo("Valor Original Corrigido", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 5: // E - Depreciação Acumulada Corrigido
                            Valida_campo("Depreciação Acumulada Corrigido", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 6: // F - Valor Original Moeda
                            Valida_campo("Valor Original Moeda", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 7: // G - Depreciação acumulada Moeda
                            Valida_campo("Depreciação acumulada Moeda", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 8: // H - Valor Original Ufir
                            Valida_campo("Valor Original Ufir", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 9: // I - Depreciação acumulada Ufir
                            Valida_campo("Depreciação acumulada Ufir", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;
                    }

                    if (columns > 9)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Legado_financeiro(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código legado documento*
                            Valida_campo("Código legado documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Número documento*
                            Valida_campo("Número documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Código da conta Solution
                            Valida_campo("Código da conta Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 4: // D - Código da conta legado*
                            Valida_campo("Código da conta legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 5: // E - Código endereço legado
                            Valida_campo("Código endereço legado", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 6: // F - Código endereço Solution
                            Valida_campo("Código endereço Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: // G - Tipo de documento*
                            List<String> dom_tipo_documento = new List<String> { "#", "C", "T", "A", "", "null", "NULL", "NULL" };
                            Valida_dominio("Tipo de documento", row[column].ToString(), rows, columns, dom_tipo_documento, true);
                            break;

                        case 8: // H - Pagamento ou recebimento*
                            List<String> dom_pagar_receber = new List<String> { "P", "R", "", "null", "NULL", "NULL" };
                            Valida_dominio("Pagamento ou recebimento", row[column].ToString(), rows, columns, dom_pagar_receber, true);
                            break;

                        case 9: // I - Código empresa Solution*
                            Valida_campo("Código empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 10: // J - Código filial Solution*
                            Valida_campo("Código filial Solution", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 11: // K - CNPJ filial
                            Valida_campo("CNPJ filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 12: // L - Data de emissão*
                            Valida_campo("Data de emissão", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 13: // M - Data de vencimento*
                            Valida_campo("Data de vencimento", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 14: // N - Portador
                            Valida_campo("Portador", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 15: // O - Número da parcela
                            Valida_campo("Número da parcela", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 16: // P - Número nota fiscal
                            Valida_campo("Número nota fiscal", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 17: // Q - Centro de custo
                            Valida_campo("Centro de custo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 18: // R - Vendedor
                            Valida_campo("Vendedor", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 19: // S - Valor*
                            Valida_campo("Valor", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 20: // T - Valor de juros
                            Valida_campo("Valor de juros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 21: // U - Valor de desconto
                            Valida_campo("Valor de desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 22: // V - Valor de multa
                            Valida_campo("Valor de multa", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 23: // W - Número febraban banco
                            Valida_campo("Número febraban banco", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 24: // X - Nosso número boleto
                            Valida_campo("Nosso número boleto", row[column].ToString(), rows, columns, "char", 30, false);
                            break;

                        case 25: // Y - Dias de atraso
                            Valida_campo("Dias de atraso", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 26: // Z - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;
                    }

                    if (columns > 26)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Legado_pagamentos(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código legado pagamento*
                            Valida_campo("Código legado pagamento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado documento*
                            Valida_campo("Código legado documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Número documento
                            Valida_campo("Número documento", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 4: // D - Código documento Solution
                            List<String> dom_codigo_documento = new List<String> { "", "null", "NULL", "NULL" };
                            Valida_dominio("Código documento Solution", row[column].ToString(), rows, columns, dom_codigo_documento, false);
                            break;

                        case 5: // E - Empresa*
                            Valida_campo("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 6: // F - CNPJ Filial
                            Valida_campo("CNPJ Filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 7: // G - Filial*
                            Valida_campo("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 8: // H - Valor*
                            Valida_campo("Valor", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 9: // I - Valor juros
                            Valida_campo("Valor juros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 10: // J - Valor multa
                            Valida_campo("Valor multa", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 11: // K - Desconto valor
                            Valida_campo("Desconto valor", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 12: // L - Data pagamento*
                            Valida_campo("Data pagamento", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;
                    }

                    if (columns > 12)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Legado_pedidos(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código pedido*
                            Valida_campo("Código pedido", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado pedido*
                            Valida_campo("Código legado pedido", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Empresa*
                            Valida_campo("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 4: // D - Filial*
                            Valida_campo("Filial", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 5: // E - CNPJ filial
                            Valida_campo("CNPJ filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 6: // F - Módulo*
                            List<String> dom_modulo = new List<String> { "5", "17", "", "null", "NULL" };
                            Valida_dominio("Módulo", row[column].ToString(), rows, columns, dom_modulo, true);
                            break;

                        case 7: // G - Tipo*
                            List<String> dom_tipo = new List<String> { "O", "P", "", "null", "NULL" };
                            Valida_dominio("Tipo", row[column].ToString(), rows, columns, dom_tipo, true);
                            break;

                        case 8: // H - Data hora abertura
                            Valida_campo("Data hora abertura", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 9: // I - Data hora validade
                            Valida_campo("Data hora validade", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 10: // J - Data hora encerramento
                            Valida_campo("Data hora encerramento", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 11: // K - Código cliente legado*
                            Valida_campo("Código cliente legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 12: // L - Código legado endereço
                            Valida_campo("Código legado endereço", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 13: // M - Código endereço Solution
                            Valida_campo("Código endereço Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 14: // N - Código cliente Solution
                            Valida_campo("Código cliente Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 15: // O - Nome cliente
                            Valida_campo("Nome cliente", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 16: // P - Logradouro cliente
                            Valida_campo("Logradouro cliente", row[column].ToString(), rows, columns, "char", 500, false);
                            break;

                        case 17: // Q - Cidade cliente
                            Valida_campo("Cidade cliente", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 18: // R - UF cliente
                            Valida_campo("UF cliente", row[column].ToString(), rows, columns, "char", 2, false);
                            break;

                        case 19: // S - CEP cliente
                            Valida_campo("CEP cliente", row[column].ToString(), rows, columns, "char", 9, false);
                            break;

                        case 20: // T - CNPJ/CPF cliente
                            Valida_campo("CNPJ/CPF cliente", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 21: // U - Inscrição estadual cliente
                            Valida_campo("Inscrição estadual cliente", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 22: // V - Inscrição municipal cliente
                            Valida_campo("Inscrição municipal cliente", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 23: // W - Vendedor
                            Valida_campo("Vendedor", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 24: // X - Politica prazo
                            Valida_campo("Politica prazo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Tipo pagamento*
                            List<String> dom_pagamento = new List<String> { "V", "P", "", "null", "NULL" };
                            Valida_dominio("Tipo pagamento", row[column].ToString(), rows, columns, dom_pagamento, true);
                            break;

                        case 26: // Z - Forma pagamento*
                            List<String> dom_forma_pagamento = new List<String> { "A", "2", "4", "5", "0", "1", "6", "3", "F", "9", "8", "", "null", "NULL" };
                            Valida_dominio("Forma pagamento", row[column].ToString(), rows, columns, dom_forma_pagamento, true);
                            break;

                        case 27: // AA - Número parcelas
                            Valida_campo("Número parcelas", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 28: // AB - Data hora parcelamento
                            Valida_campo("Data hora parcelamento", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 29: // AC - Operação
                            Valida_campo("Operação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 30: // AD - Número nota fiscal
                            Valida_campo("Número nota fiscal", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 31: // AE - Chave nota fiscal
                            Valida_campo("Chave nota fiscal", row[column].ToString(), rows, columns, "char", 50, false);
                            break;

                        case 32: // AF - Valor de outras despesas
                            Valida_campo("Valor de outras despesas", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 33: // AG - Valor frete
                            Valida_campo("Valor frete", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 34: // AH - Valor desconto
                            Valida_campo("Valor desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 35: // AI - Valor impostos adicionais
                            Valida_campo("Valor impostos adicionais", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 36: // AJ - Valor total*
                            Valida_campo("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 37: // AK - Código veículo Solution
                            Valida_campo("Código veículo Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 38: // AL - Código veículo legado
                            Valida_campo("Código veículo legado", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 39: // AM - Número serie veículo
                            Valida_campo("Número serie veículo", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 40: // AN - Classificação
                            Valida_campo("Classificação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 41: // AO - Hodometro
                            Valida_campo("Hodometro", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 42: // AP - Horimetro
                            Valida_campo("Horimetro", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 43: // AQ - Mecanico
                            Valida_campo("Mecanico", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 44: // AR - Tipo ordem serviço
                            Valida_campo("Tipo ordem serviço", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 45: // AS - Descrição problema
                            Valida_campo("Descrição problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 46: // AT - Opinião do problema
                            Valida_campo("Opinião do problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 47: // AU - Solução problema
                            Valida_campo("Solução problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 48: // AV - Total km rodados
                            Valida_campo("Total km rodados", row[column].ToString(), rows, columns, "numeric", 16.1, false);
                            break;

                        case 49: // AW - Total valor deslocamento
                            Valida_campo("Total valor deslocamento", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 50: // AX - Total valor KM
                            Valida_campo("Total valor KM", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 51: // AY - Total valor serviços
                            Valida_campo("Total valor serviços", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 52: // AZ - Total valor serviço de terceiros
                            Valida_campo("Total valor serviço de terceiros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 53: // BA - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;
                    }

                    if (columns > 53)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Legado_pedidos_itens(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código item*
                            Valida_campo("Código item", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado item*
                            Valida_campo("Código legado item", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Código legado pedido*
                            Valida_campo("Código legado pedido", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 4: // D - Código pedido Solution
                            Valida_campo("Código pedido Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 5: // E - Empresa*
                            Valida_campo("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 6: // F - Tipo item*
                            List<String> dom_tipo_item = new List<String> { "SP", "P", "ST", "", "null", "NULL" };
                            Valida_dominio("Tipo item", row[column].ToString(), rows, columns, dom_tipo_item, true);
                            break;

                        case 7: // G - Código produto Solution
                            Valida_campo("Código produto Solution", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 8: // H - Código produto legado*
                            Valida_campo("Código produto legado", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 9: // I - Descrição produto
                            Valida_campo("Descrição produto", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 10: // J - Data hora alocação
                            Valida_campo("Data hora alocação", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 11: // K - Unidade
                            Valida_campo("Unidade", row[column].ToString(), rows, columns, "char", 6, false);
                            break;

                        case 12: // L - Código item pedido fornecedor
                            Valida_campo("Código item pedido fornecedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 13: // M - Número pedido fornecedor
                            Valida_campo("Número pedido fornecedor", row[column].ToString(), rows, columns, "char", 15, false);
                            break;

                        case 14: // N - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 15: // O - Preço unitário
                            Valida_campo("Preço unitário", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 16: // P - Valor desconto
                            Valida_campo("Valor desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 17: // Q - Valor frete
                            Valida_campo("Valor frete", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 18: // R - Valor impostos adicionais
                            Valida_campo("Valor impostos adicionais", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 19: // S - Valor outras despesas
                            Valida_campo("Valor outras despesas", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 20: // T - Valor total*
                            Valida_campo("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 21: // U - Tipo calculo
                            Valida_campo("Tipo calculo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 22: // V - Total horas trabalhadas
                            Valida_campo("Total horas trabalhadas", row[column].ToString(), rows, columns, "numeric", 16.8, false);
                            break;

                        case 23: // W - Total horas vendidas
                            Valida_campo("Total horas vendidas", row[column].ToString(), rows, columns, "numeric", 16.8, false);
                            break;

                        case 24: // X - Mecanico
                            Valida_campo("Mecanico", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Observação
                            Valida_campo("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;
                    }

                    if (columns > 25)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
        }

        public void Legado_movimentacao(DataTable dataTable, int rows)
        {

            Progresso(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código empresa Solution*
                            Valida_campo("Código empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código filial Solution*
                            Valida_campo("Código filial Solution", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 3: // C - CNPJ Filial
                            Valida_campo("CNPJ Filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 4: // D - Código produto Solution
                            Valida_campo("Código produto Solution", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 5: // E - Código produto legado*
                            Valida_campo("Código produto legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 6: // F - Grupo/classificação produto
                            Valida_campo("Grupo/classificação produto", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 7: // G - Operação
                            Valida_campo("Operação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 8: // H - Tipo movimentação*
                            List<String> dom_tipo_movimentacao = new List<String> { "S", "E", "", "null", "NULL" };
                            Valida_dominio("Tipo movimentação", row[column].ToString(), rows, columns, dom_tipo_movimentacao, true);
                            break;

                        case 9: // I - Movimenta estoque*
                            List<String> dom_movimenta_estoque = new List<String> { "S", "N", "", "null", "NULL" };
                            Valida_dominio("Movimenta estoque", row[column].ToString(), rows, columns, dom_movimenta_estoque, true);
                            break;

                        case 10: // J - Número documento
                            Valida_campo("Número documento", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 11: // K - Data movimentação
                            Valida_campo("Data movimentação", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 12: // L - hora movimentação
                            Valida_campo("hora movimentação", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 13: // M - Quantidade*
                            Valida_campo("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 14: // N - Custo médio total
                            Valida_campo("Custo médio total", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 15: // O - Valor total*
                            Valida_campo("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;
                    }

                    if (columns > 15)
                    {
                        Adicionar_registro("Erro genérico", rows, columns, row[column].ToString(), "Excedeu o número de colunas");
                    }

                    columns++;
                }

                Atualizar_progresso(total, rows);

                rows++;
            }

            Progresso(false);
            Grid_criar();
        }
    }
}
