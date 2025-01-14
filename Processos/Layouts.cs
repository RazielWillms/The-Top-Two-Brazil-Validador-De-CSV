using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using static ValidarCSV.TypeExtensions;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        private void Validar_layouts_gerenciar(DataTable dataTable, String Tabela)
        {
            int rows = 0;

            if (Cabecalho.SelectedIndex == Indice_Cabecalho_Retornar(CabecalhoType.Sim))
            {
                rows = 1;
            }

            LayoutType layout = LayoutType.Indefinido;

            Layout_enum_retornar(Tabela, ref layout);

            switch (layout)
            {
                case LayoutType.Maquinas:
                    Maquinas(dataTable, rows);
                    break;

                case LayoutType.MaquinasCompleto:
                    MaquinasCompleto(dataTable, rows);
                    break;

                case LayoutType.Produtos:
                    Produtos(dataTable, rows);
                    break;

                case LayoutType.SaldosMaquinas:
                    Saldos_maquinas(dataTable, rows);
                    break;

                case LayoutType.Adiantamentos:
                    Adiantamentos(dataTable, rows);
                    break;

                case LayoutType.OrcamentoBalcao:
                    Orcamento_balcao(dataTable, rows);
                    break;

                case LayoutType.OrcamentoOficina:
                    Orcamento_oficina(dataTable, rows);
                    break;

                case LayoutType.Estatisticas:
                    Estatisticas(dataTable, rows);
                    break;

                case LayoutType.VeiculosClientes:
                    Veiculos_clientes(dataTable, rows);
                    break;

                case LayoutType.ImobilizadoItens:
                    Imobilizado_itens(dataTable, rows);
                    break;

                case LayoutType.ImobilizadoSaldos:
                    Imobilizado_saldos(dataTable, rows);
                    break;

                case LayoutType.LegadoFinanceiro:
                    Legado_financeiro(dataTable, rows);
                    break;

                case LayoutType.LegadoPagamentos:
                    Legado_pagamentos(dataTable, rows);
                    break;

                case LayoutType.LegadoPedidos:
                    Legado_pedidos(dataTable, rows);
                    break;

                case LayoutType.LegadoPedidosItens:
                    Legado_pedidos_itens(dataTable, rows);
                    break;

                case LayoutType.LegadoMovimentacao:
                    Legado_movimentacao(dataTable, rows);
                    break;

                case LayoutType.Grupos:
                    Grupos(dataTable, rows);
                    break;

                case LayoutType.SubGrupos:
                    Sub_grupos(dataTable, rows);
                    break;

                case LayoutType.Plano:
                    Plano_de_Contas(dataTable, rows);
                    break;

                case LayoutType.Contas:
                    Contas(dataTable, rows);
                    break;

                case LayoutType.Indefinido:
                    MessageBox.Show("A validação deste layout ainda não foi implementada", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;

                default:
                    MessageBox.Show("A validação deste layout ainda não foi implementada", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }
        }

        public void Maquinas(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código do Produto*
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 2: //B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: //C - Descrição adicional do item*
                            Campos_validar_gerenciar("Descrição adicional do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, true);
                            break;

                        case 4: //D - Tipo de mercadoria*
                            Campos_validar_gerenciar("Tipo de mercadoria", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, true);
                            break;

                        case 5: //E - Marca
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 6: //F - Departamento
                            Campos_validar_gerenciar("Departamento", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 7: //G - Controla estoque
                            Campos_validar_gerenciar("Controla estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controla_estoque), false);
                            break;

                        case 8: //H - Código do grupo*
                            Campos_validar_gerenciar("Código do grupo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, true);
                            break;

                        case 9: //I - Peso liquido
                            Campos_validar_gerenciar("Peso Liquido", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 10: //J - Peso bruto
                            Campos_validar_gerenciar("Peso bruto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 11: //K - Unidade*
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, true);
                            break;

                        case 12: //L - Aplicação
                            Campos_validar_gerenciar("Aplicação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 13: //M - Apelido
                            Campos_validar_gerenciar("Apelido", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 14: //N - Produto Importado ou Nacional
                            Campos_validar_gerenciar("Importado ou Nacional", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Importado_nacional), false);
                            break;

                        case 15: //O - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 16: //P - Preço de reposição
                            Campos_validar_gerenciar("Preço de reposição", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 17: //Q - Código de referência
                            Campos_validar_gerenciar("Código de referência", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 18: //R - Situação
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Situacao), false);
                            break;

                        case 19: //S - Produto usado*
                            Campos_validar_gerenciar("Produto usado", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Usado), true);
                            break;

                        case 20: //T - NCM*
                            Campos_validar_gerenciar("NCM", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, true);
                            break;

                        case 21: //U - Modelo
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 22: //V - Classe produto*
                            Campos_validar_gerenciar("Classe produto", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Classe), true);
                            break;

                        case 23: //W - Código base*
                            Campos_validar_gerenciar("Código base", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 24: //X - Número de serie
                            Campos_validar_gerenciar("Número de serie", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 25: //Y - Código Fiscal
                            Campos_validar_gerenciar("Código Fiscal", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 26: //Z - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 27: //AA - Controle de estoque*
                            Campos_validar_gerenciar("Controle de estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controle), true);
                            break;

                        case 28: //AB - Campo Livre
                            Campos_validar_gerenciar("Campo Livre", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 29: //AC - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 30: //AD - Código bandeira*
                            Campos_validar_gerenciar("Código bandeira", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, true);
                            break;

                        case 31: //AE - FINAME
                            Campos_validar_gerenciar("FINAME", row[column].ToString(), rows, columns, TipoCampoType.Integer, 60, false);
                            break;

                        case 32: //AG - Ano Fabricação
                            Campos_validar_gerenciar("Ano Fabricação", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;
                    }

                    if (columns > 32)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }
                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void MaquinasCompleto(DataTable dataTable, int rows)
        {
            // Dicionário para armazenar o código do item e a linha onde ele apareceu pela primeira vez
            Dictionary<string, string> listaCodItem = new Dictionary<string, string>();
            string codigoItem = string.Empty;
            string classeItem = string.Empty;

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código do Produto*
                            codigoItem = row[column].ToString();
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 2: //B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: //C - Descrição adicional do item*
                            Campos_validar_gerenciar("Descrição adicional do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, true);
                            break;

                        case 4: //D - Tipo de mercadoria*
                            Campos_validar_gerenciar("Tipo de mercadoria", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, true);
                            break;

                        case 5: //E - Marca
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 6: //F - Departamento
                            Campos_validar_gerenciar("Departamento", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 7: //G - Controla estoque
                            Campos_validar_gerenciar("Controla estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controla_estoque), false);
                            break;

                        case 8: //H - Código do grupo*
                            Campos_validar_gerenciar("Código do grupo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, true);
                            break;

                        case 9: //I - Peso liquido
                            Campos_validar_gerenciar("Peso Liquido", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 10: //J - Peso bruto
                            Campos_validar_gerenciar("Peso bruto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 11: //K - Unidade*
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, true);
                            break;

                        case 12: //L - Aplicação
                            Campos_validar_gerenciar("Aplicação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 13: //M - Apelido
                            Campos_validar_gerenciar("Apelido", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 14: //N - Produto Importado ou Nacional
                            Campos_validar_gerenciar("Importado ou Nacional", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Importado_nacional), false);
                            break;

                        case 15: //O - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 16: //P - Preço de reposição
                            Campos_validar_gerenciar("Preço de reposição", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 17: //Q - Código de referência
                            Campos_validar_gerenciar("Código de referência", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 18: //R - Situação
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Situacao), false);
                            break;

                        case 19: //S - Produto usado*
                            Campos_validar_gerenciar("Produto usado", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Usado), true);
                            break;

                        case 20: //T - NCM*
                            Campos_validar_gerenciar("NCM", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, true);
                            break;

                        case 21: //U - Modelo
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 22: //V - Classe produto*
                            classeItem = row[column].ToString();

                            // Verifica se o código base já existe, se não, adiciona
                            if (!listaCodItem.ContainsKey(codigoItem))
                            {
                                listaCodItem[codigoItem] = classeItem;
                            }
                            Campos_validar_gerenciar("Classe produto", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Classe), true);
                            break;

                        case 23: //W - Código base*
                            if (!listaCodItem.ContainsKey(row[column].ToString()) || listaCodItem[row[column].ToString()].ToUpper() != "B")
                            {
                                Registro_adicionar("Código base", rows, columns, row[column].ToString(), $"Cadastro base não existe ou deve ser inforamdo antes do individual");
                            }

                            Campos_validar_gerenciar("Código base", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 24: //X - Número de serie
                            Campos_validar_gerenciar("Número de serie", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 25: //Y - Código Fiscal
                            Campos_validar_gerenciar("Código Fiscal", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 26: //Z - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 27: //AA - Controle de estoque*
                            Campos_validar_gerenciar("Controle de estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controle), true);
                            break;

                        case 28: //AB - Campo Livre
                            Campos_validar_gerenciar("Campo Livre", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 29: //AC - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 30: //AD - Código bandeira*
                            Campos_validar_gerenciar("Código bandeira", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, true);
                            break;

                        case 31: //AE - FINAME
                            Campos_validar_gerenciar("FINAME", row[column].ToString(), rows, columns, TipoCampoType.Integer, 60, false);
                            break;

                        case 32: //AG - Ano Fabricação
                            Campos_validar_gerenciar("Ano Fabricação", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;
                    }

                    if (columns > 32)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }
                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Saldos_maquinas(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: //B - Código do Produto*
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 3: //C - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.4, true);
                            break;

                        case 4: //D - Valor do Estoque*
                            Campos_validar_gerenciar("Valor do Estoque", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, true);
                            break;

                        case 5: //E - Código da prateleira
                            Campos_validar_gerenciar("Código da prateleira", row[column].ToString(), rows, columns, TipoCampoType.Character, 14, false);
                            break;

                        case 6: //F - Data da última compra
                            Campos_validar_gerenciar("Data da última compra", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 7: //G - Valor da última compra
                            Campos_validar_gerenciar("Valor da última compra", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 8: //H - Estoque mínimo
                            Campos_validar_gerenciar("Estoque mínimo", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 9: //I - Descrição
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 10: //J - Código produto único
                            Campos_validar_gerenciar("Código produto único", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 11: //K - Custo Reposição
                            Campos_validar_gerenciar("Estoque mínimo", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 15.2, false);
                            break;

                        case 12: //L - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 15.3, false);
                            break;
                    }

                    if (columns > 12)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Adiantamentos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: //B - Conta legado*
                            Campos_validar_gerenciar("Conta legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 3: //C - Valor do adiantamento*
                            Campos_validar_gerenciar("Valor do adiantamento", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;

                        case 4: //D - Tipo do adiantamento*
                            Campos_validar_gerenciar("Tipo do adiantamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Cliente_fornecedor), true);
                            break;

                        case 5: //E - Centro de Custo
                            Campos_validar_gerenciar("Centro de Custo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 6: //F - Número
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 7: //G - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;
                    }

                    if (columns > 7)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Orcamento_balcao(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código Pedido*
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, true);
                            break;

                        case 2: //B - Código do cliente (sistema antigo)*
                            Campos_validar_gerenciar("Código Legado do Cliente", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 3: //C - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, TipoCampoType.Integer, 3, false);
                            break;

                        case 4: //D - Política Prazo
                            Campos_validar_gerenciar("Política Prazo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 5: //E - Politica Preço
                            Campos_validar_gerenciar("Politica Preço", row[column].ToString(), rows, columns, TipoCampoType.Integer, 3, false);
                            break;

                        case 6: //F - Tipo Operação
                            Campos_validar_gerenciar("Tipo Operação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_operacao), false);
                            break;

                        case 7: //G - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 8: //H - Funcionário Abertura O.C
                            Campos_validar_gerenciar("Funcionário Abertura O.C", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 9: //I - Data Validade
                            Campos_validar_gerenciar("Data Validade", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 10: //J - Data Abertura*
                            Campos_validar_gerenciar("Data Abertura", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, true);
                            break;

                        case 11: //K - Data Parcelamento
                            Campos_validar_gerenciar("Data Parcelamento", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 12: //L - Situação*
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Orcamento_situacao), false);
                            break;

                        case 13: //M - Status*
                            Campos_validar_gerenciar("Status", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Status), true);
                            break;

                        case 14: //N - Produto*
                            Campos_validar_gerenciar("Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 15: //O - Descrição Produto
                            Campos_validar_gerenciar("Descrição Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 16: //P - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, true);
                            break;

                        case 17: //Q - Preço Unitário*
                            Campos_validar_gerenciar("Preço Unitário", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.3, true);
                            break;

                        case 18: //R - Valor Desconto
                            Campos_validar_gerenciar("Valor Desconto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 19: //S - Vendedor Produto
                            Campos_validar_gerenciar("Vendedor Produto", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;
                    }

                    if (columns > 19)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Orcamento_oficina(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                //variáveis para analisar excessão
                string serieVeiculo = null;
                string veiculoId = null;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: //A - Número*
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, true);
                            break;

                        case 2: //B - Código da Filial Solution* (pode ser 0, por isto false no obrigatório)
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 3: //C - ID do Veículo*
                            veiculoId = row[column].ToString();
                            Campos_validar_gerenciar("ID do Veículo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 4: //D - Série do veículo*
                            serieVeiculo = row[column].ToString();
                            Campos_validar_gerenciar("Série do veículo", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);

                            //excessão de tratamento
                            if (string.IsNullOrWhiteSpace(veiculoId) && string.IsNullOrWhiteSpace(serieVeiculo))
                            {
                                string mensagem = "Informar ID do Veículo ou sua Série";
                                Registro_adicionar("Erro Genérico", rows, 0, null, mensagem);
                            }
                            break;

                        case 5: //E - Conta do cliente legado - sistema antigo*
                            Campos_validar_gerenciar("Conta do cliente legado", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, true);
                            break;

                        case 6: //F - Tipo da OS
                            Campos_validar_gerenciar("Tipo da OS", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 7: //G - Data de abertura
                            Campos_validar_gerenciar("Data de abertura", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 8: //H - ID do mecânico no Solution
                            Campos_validar_gerenciar("Mecânico no Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 9: //I - ID do vendedor no Solution
                            Campos_validar_gerenciar("Vendedor no Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 10: //J - ID do local de venda
                            Campos_validar_gerenciar("local de venda", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 11: //K - ID da política de preço
                            Campos_validar_gerenciar("política de preço", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 12: //L - ID da política de prazo
                            Campos_validar_gerenciar("política de prazo", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 13: //M - Código do produto*
                            Campos_validar_gerenciar("Código do produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 14: //N - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, true);
                            break;

                        case 15: //O - Preço unitário*
                            Campos_validar_gerenciar("Preço unitário", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;
                    }

                    if (columns > 15)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Estatisticas(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código filial Solution*
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: //B - Código produto*
                            Campos_validar_gerenciar("Código produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 3: //C - Data movimetação (mês e ano)*
                            Campos_validar_gerenciar("Data movimetação", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, true);
                            break;

                        case 4: //D - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 15.4, true);
                            break;

                        case 5: //E - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;
                    }

                    if (columns > 5)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Veiculos_clientes(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código*
                            Campos_validar_gerenciar("Código", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, true);
                            break;

                        case 2: // B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: // C - Placa
                            Campos_validar_gerenciar("Placa", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 4: // D - Meses Garantia
                            Campos_validar_gerenciar("Meses Garantia", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 5: // E - Hrs.Garantia
                            Campos_validar_gerenciar("Hrs.Garantia", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 6: // F - Km garantia
                            Campos_validar_gerenciar("Km garantia", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 10.1, false);
                            break;

                        case 7: // G - Novo Usado*
                            Campos_validar_gerenciar("Novo Usado", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Novo_usado), true);
                            break;

                        case 8: // H - Versão
                            Campos_validar_gerenciar("Versão", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 9: // I - Ano fabricação*
                            Campos_validar_gerenciar("Ano fabricação", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 10: // J - Ano modelo*
                            Campos_validar_gerenciar("Ano modelo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 11: // K - Código da conta de cliente (sistema antigo)*
                            Campos_validar_gerenciar("Código da conta de cliente (sistema antigo)", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, true);
                            break;

                        case 12: // L - Modelo*
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, true);
                            break;

                        case 13: // M - numero NF de compra
                            Campos_validar_gerenciar("numero NF de compra", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 14: // N - Data de compra
                            Campos_validar_gerenciar("Data de compra", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 15: // O - Código da conta de fornecedor
                            Campos_validar_gerenciar("Código da conta de fornecedor", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 16: // P - Nome fornecedor
                            Campos_validar_gerenciar("Nome fornecedor", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 17: // Q - Código produto estoque
                            Campos_validar_gerenciar("Código produto estoque", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 18: // R - Numero de serie*
                            Campos_validar_gerenciar("Numero de serie", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 19: // S - Serie motor*
                            Campos_validar_gerenciar("Serie motor", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, true);
                            break;

                        case 20: // T - Série da bomba hidráulica
                            Campos_validar_gerenciar("Série da bomba hidráulica", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 21: // U - Série de transmissão
                            Campos_validar_gerenciar("Série de transmissão", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 22: // V - Série da caixa de câmbio
                            Campos_validar_gerenciar("Série da caixa de câmbio", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 23: // W - Série da bomba injetora
                            Campos_validar_gerenciar("Série da bomba injetora", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 24: // X - Série do monobloco
                            Campos_validar_gerenciar("Série do monobloco", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 25: // Y - Série do eixo dianteiro
                            Campos_validar_gerenciar("Série do eixo dianteiro", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 26: // Z - Série da plataforma
                            Campos_validar_gerenciar("Série da plataforma", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 27: // AA - Pneus dianteiro
                            Campos_validar_gerenciar("Pneus dianteiro", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 28: // AB - Pneus traseiro
                            Campos_validar_gerenciar("Pneus traseiro", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 29: // AC - Série direção hidráulica
                            Campos_validar_gerenciar("Série direção hidráulica", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 30: // AD - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 200, false);
                            break;

                        case 31: // AE - Tipo equipamento*
                            Campos_validar_gerenciar("Tipo equipamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_equipamento), true);
                            break;

                        case 32: // AF - Código do pedido da gestão de compra
                            Campos_validar_gerenciar("Código do pedido da gestão de compra", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 33: // AG - Cor código*
                            Campos_validar_gerenciar("Cor código", row[column].ToString(), rows, columns, TipoCampoType.Character, 4, true);
                            break;

                        case 34: // AH - Cor descrição*
                            Campos_validar_gerenciar("Cor descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 35: // AI - Potência do Motor (CV)
                            Campos_validar_gerenciar("Potência do Motor (CV)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 8.1, false);
                            break;

                        case 36: // AJ - CM3 (cilindradas)
                            Campos_validar_gerenciar("CM3 (cilindradas)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 8.1, false);
                            break;

                        case 37: // AK - Peso líquido (KG)
                            Campos_validar_gerenciar("Peso líquido (KG)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 38: // AL - Peso bruto (KG)
                            Campos_validar_gerenciar("Peso bruto (KG)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 39: // AM - Tipo combustivel*
                            Campos_validar_gerenciar("Tipo combustivel", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, true);
                            break;

                        case 40: // AN - CMKG
                            Campos_validar_gerenciar("CMKG", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 41: // AO - TMA
                            Campos_validar_gerenciar("TMA", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 42: // AP - Distância entre eixos (mm)
                            Campos_validar_gerenciar("Distância entre eixos (mm)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 8.2, false);
                            break;

                        case 43: // AQ - RENAVAM
                            Campos_validar_gerenciar("RENAVAM", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 44: // AR - Tipo pintura*
                            Campos_validar_gerenciar("Tipo pintura", row[column].ToString(), rows, columns, TipoCampoType.Character, 1, true);
                            break;

                        case 45: // AS - Tipo de Veículo Renavam/Denatran
                            Campos_validar_gerenciar("Tipo de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_renavam_denatram), false);
                            break;

                        case 46: // AT - Espécie de Veículo Renavam/Denatran
                            Campos_validar_gerenciar("Espécie de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Especie_veiculo_renavam_denatram), false);
                            break;

                        case 47: // AU - Marca Modelo Renavam/Denatran
                            Campos_validar_gerenciar("Marca Modelo Renavam/Denatran", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 48: // AV - Codigo do DN
                            Campos_validar_gerenciar("Codigo do DN", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 49: // AW - Chassis*
                            Campos_validar_gerenciar("Chassis", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, true);
                            break;

                        case 50: // AX - Marca
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 51: // AY - Data entrega tecnica
                            Campos_validar_gerenciar("Data entrega tecnica", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 52: // AZ - Data ultima revisão
                            Campos_validar_gerenciar("Data ultima revisão", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;
                    }

                    if (columns > 52)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Imobilizado_itens(DataTable dataTable, int rows)
        {
            // Dicionário para armazenar o código do item e a linha onde ele apareceu pela primeira vez
            Dictionary<string, int> listaCodItem = new Dictionary<string, int>();

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código da Empresa Solution*
                            Campos_validar_gerenciar("Código da Empresa Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: // B - Código da Filial Solution*
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 3: // C - Código do Item*

                            string codigoItem = row[column].ToString();

                            // Verifica se o código do item já existe
                            if (!listaCodItem.ContainsKey(codigoItem))
                            {
                                // Armazena o código do item e a linha atual
                                listaCodItem[codigoItem] = rows;
                            }
                            else
                            {
                                // Recupera a linha original onde o código foi encontrado
                                int linhaOriginal = listaCodItem[codigoItem];
                                Registro_adicionar("Código do Item", rows, columns, codigoItem, $"Não é permitido repetir códigos entre itens, confira as linhas: {linhaOriginal} e {rows}");
                            }

                            Campos_validar_gerenciar("Código do Item", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, true);
                            break;

                        case 4: // D - Código da Conta (Plano de Contas)
                            Campos_validar_gerenciar("Código da Conta (Plano de Contas)", row[column].ToString(), rows, columns, TipoCampoType.Character, 11, true);
                            break;

                        case 5: // E - Data do lancto*
                            Campos_validar_gerenciar("Data do lancto", row[column].ToString(), rows, columns, TipoCampoType.Date, 10, true);
                            break;

                        case 6: // F - Data da aquisição*
                            Campos_validar_gerenciar("Data da aquisição", row[column].ToString(), rows, columns, TipoCampoType.Date, 10, true);
                            break;

                        case 7: // G - Centro de Custo*
                            Campos_validar_gerenciar("Centro de Custo", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, true);
                            break;

                        case 8: // H - % de Depreciação do Item
                            Campos_validar_gerenciar("% de Depreciação do Item", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 9: // I - % de Depreciação Gerencial
                            Campos_validar_gerenciar("% de Depreciação Gerencial", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 10: // J - % residual
                            Campos_validar_gerenciar("% residual", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 5.2, false);
                            break;

                        case 11: // K - Débito ou Crédito*
                            Campos_validar_gerenciar("Débito ou Crédito", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Debito_credito), true);
                            break;

                        case 12: // L - Chave*
                            Campos_validar_gerenciar("Chave", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Chave), true);
                            break;

                        case 13: // M - Tipo lançamento
                            Campos_validar_gerenciar("Tipo lançamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_lancamento), true);
                            break;

                        case 14: // N - Tipo Baixa
                            Campos_validar_gerenciar("Tipo lançamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_baixa), false);
                            break;

                        case 15: // O - Número do documento de aquisição
                            Campos_validar_gerenciar("Número do documento de aquisição", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 16: // P - Nome do Fornecedor
                            Campos_validar_gerenciar("Nome do Fornecedor", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 17: // Q - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 225, true);
                            break;

                        case 18: // R - Descrição sucienta da função do bem na atividade do estabelecimento (obrigatório para Sped Fiscal)*
                            Campos_validar_gerenciar("Descrição sucienta", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 19: // S - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 20: // T - Número da Apólice
                            Campos_validar_gerenciar("Número da Apólice", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 21: // U - Data do Vencimento
                            Campos_validar_gerenciar("Data do Vencimento", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 22: // V - Código Externo
                            Campos_validar_gerenciar("Código Externo", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 23: // W - Código do Local
                            Campos_validar_gerenciar("Código do Local", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 24: // X - Código do Responsável
                            Campos_validar_gerenciar("Código do Responsável", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 25: // Y - Código do tipo do bem
                            Campos_validar_gerenciar("Código do tipo do bem", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 26: // Z - Código da Seguradora
                            Campos_validar_gerenciar("Código da Seguradora", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 27: // AA - Tipo Documento de aquisição
                            Campos_validar_gerenciar("Tipo Documento de aquisição", row[column].ToString(), rows, columns, TipoCampoType.Integer, 3, false);
                            break;

                        case 28: // AB - Situação do Bem
                            Campos_validar_gerenciar("Situação do Bem", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 29: // AC - Chassis
                            Campos_validar_gerenciar("Chassis", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, false);
                            break;

                        case 30: // AD - Placa
                            Campos_validar_gerenciar("Placa", row[column].ToString(), rows, columns, TipoCampoType.Character, 9, false);
                            break;
                    }

                    if (columns > 30)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Imobilizado_saldos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código da Empresa*
                            Campos_validar_gerenciar("Código da Empresa", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: // B - Código do Item*
                            Campos_validar_gerenciar("Código do Item", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, true);
                            break;

                        case 3: // C - Valor Original* //a princípio era obrigatório, mas não tem validação na conversão e cliente reclamou de gerar erro quando informado '0'
                            Campos_validar_gerenciar("Valor Original", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 15.2, false);
                            break;

                        case 4: // D - Valor Original Corrigido
                            Campos_validar_gerenciar("Valor Original Corrigido", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 5: // E - Depreciação Acumulada Corrigido
                            Campos_validar_gerenciar("Depreciação Acumulada Corrigido", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 6: // F - Valor Original Moeda
                            Campos_validar_gerenciar("Valor Original Moeda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 7: // G - Depreciação acumulada Moeda
                            Campos_validar_gerenciar("Depreciação acumulada Moeda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 8: // H - Valor Original Ufir
                            Campos_validar_gerenciar("Valor Original Ufir", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;

                        case 9: // I - Depreciação acumulada Ufir
                            Campos_validar_gerenciar("Depreciação acumulada Ufir", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, false);
                            break;
                    }

                    if (columns > 9)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Legado_financeiro(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código legado documento*
                            Campos_validar_gerenciar("Código legado documento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 2: // B - Número documento*
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 3: // C - Código da conta Solution
                            Campos_validar_gerenciar("Código da conta Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 4: // D - Código da conta legado*
                            Campos_validar_gerenciar("Código da conta legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 5: // E - Código endereço legado
                            Campos_validar_gerenciar("Código endereço legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 6: // F - Código endereço Solution
                            Campos_validar_gerenciar("Código endereço Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 7: // G - Tipo de documento*
                            Campos_validar_gerenciar("Tipo de documento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_documento), true);
                            break;

                        case 8: // H - Pagamento ou recebimento*
                            Campos_validar_gerenciar("Pagamento ou recebimento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Pagar_receber), true);
                            break;

                        case 9: // I - Código empresa Solution*
                            Campos_validar_gerenciar("Código empresa Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 10: // J - Código filial Solution*
                            Campos_validar_gerenciar("Código filial Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 2, true);
                            break;

                        case 11: // K - CNPJ filial
                            Campos_validar_gerenciar("CNPJ filial", row[column].ToString(), rows, columns, TipoCampoType.Character, 18, false);
                            break;

                        case 12: // L - Data de emissão*
                            Campos_validar_gerenciar("Data de emissão", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 13: // M - Data de vencimento*
                            Campos_validar_gerenciar("Data de vencimento", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 14: // N - Portador
                            Campos_validar_gerenciar("Portador", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 15: // O - Número da parcela
                            Campos_validar_gerenciar("Número da parcela", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 16: // P - Número nota fiscal
                            Campos_validar_gerenciar("Número nota fiscal", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 17: // Q - Centro de custo
                            Campos_validar_gerenciar("Centro de custo", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 18: // R - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 19: // S - Valor*
                            Campos_validar_gerenciar("Valor", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;

                        case 20: // T - Valor de juros
                            Campos_validar_gerenciar("Valor de juros", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 21: // U - Valor de desconto
                            Campos_validar_gerenciar("Valor de desconto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 22: // V - Valor de multa
                            Campos_validar_gerenciar("Valor de multa", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 23: // W - Número febraban banco
                            Campos_validar_gerenciar("Número febraban banco", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 24: // X - Nosso número boleto
                            Campos_validar_gerenciar("Nosso número boleto", row[column].ToString(), rows, columns, TipoCampoType.Character, 30, false);
                            break;

                        case 25: // Y - Dias de atraso
                            Campos_validar_gerenciar("Dias de atraso", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, false);
                            break;

                        case 26: // Z - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;
                    }

                    if (columns > 26)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Legado_pagamentos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código legado pagamento*
                            Campos_validar_gerenciar("Código legado pagamento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 2: // B - Código legado documento*
                            Campos_validar_gerenciar("Código legado documento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 3: // C - Número documento
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 4: // D - Código documento Solution
                            Campos_validar_gerenciar("Código documento Solution", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Null), false);
                            break;

                        case 5: // E - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 6: // F - CNPJ Filial
                            Campos_validar_gerenciar("CNPJ Filial", row[column].ToString(), rows, columns, TipoCampoType.Character, 18, false);
                            break;

                        case 7: // G - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 8: // H - Valor*
                            Campos_validar_gerenciar("Valor", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;

                        case 9: // I - Valor juros
                            Campos_validar_gerenciar("Valor juros", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 10: // J - Valor multa
                            Campos_validar_gerenciar("Valor multa", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 11: // K - Desconto valor
                            Campos_validar_gerenciar("Desconto valor", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 12: // L - Data pagamento*
                            Campos_validar_gerenciar("Data pagamento", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;
                    }

                    if (columns > 12)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Legado_pedidos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código pedido*
                            Campos_validar_gerenciar("Código pedido", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 2: // B - Código legado pedido*
                            Campos_validar_gerenciar("Código legado pedido", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 3: // C - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 4: // D - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, TipoCampoType.Integer, 2, true);
                            break;

                        case 5: // E - CNPJ filial
                            Campos_validar_gerenciar("CNPJ filial", row[column].ToString(), rows, columns, TipoCampoType.Character, 18, false);
                            break;

                        case 6: // F - Módulo*
                            Campos_validar_gerenciar("Módulo", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Modulo), true);
                            break;

                        case 7: // G - Tipo*
                            Campos_validar_gerenciar("Tipo", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo), true);
                            break;

                        case 8: // H - Data hora abertura
                            Campos_validar_gerenciar("Data hora abertura", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 9: // I - Data hora validade
                            Campos_validar_gerenciar("Data hora validade", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 10: // J - Data hora encerramento
                            Campos_validar_gerenciar("Data hora encerramento", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 11: // K - Código cliente legado*
                            Campos_validar_gerenciar("Código cliente legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 12: // L - Código legado endereço
                            Campos_validar_gerenciar("Código legado endereço", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 13: // M - Código endereço Solution
                            Campos_validar_gerenciar("Código endereço Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 14: // N - Código cliente Solution
                            Campos_validar_gerenciar("Código cliente Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 15: // O - Nome cliente
                            Campos_validar_gerenciar("Nome cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 16: // P - Logradouro cliente
                            Campos_validar_gerenciar("Logradouro cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 500, false);
                            break;

                        case 17: // Q - Cidade cliente
                            Campos_validar_gerenciar("Cidade cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 18: // R - UF cliente
                            Campos_validar_gerenciar("UF cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 2, false);
                            break;

                        case 19: // S - CEP cliente
                            Campos_validar_gerenciar("CEP cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 9, false);
                            break;

                        case 20: // T - CNPJ/CPF cliente
                            Campos_validar_gerenciar("CNPJ/CPF cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 18, false);
                            break;

                        case 21: // U - Inscrição estadual cliente
                            Campos_validar_gerenciar("Inscrição estadual cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 22: // V - Inscrição municipal cliente
                            Campos_validar_gerenciar("Inscrição municipal cliente", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 23: // W - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 24: // X - Politica prazo
                            Campos_validar_gerenciar("Politica prazo", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 25: // Y - Tipo pagamento*
                            Campos_validar_gerenciar("Tipo pagamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Pagamento), true);
                            break;

                        case 26: // Z - Forma pagamento*
                            Campos_validar_gerenciar("Forma pagamento", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Forma_pagamento), true);
                            break;

                        case 27: // AA - Número parcelas
                            Campos_validar_gerenciar("Número parcelas", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 28: // AB - Data hora parcelamento
                            Campos_validar_gerenciar("Data hora parcelamento", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 29: // AC - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 30: // AD - Número nota fiscal
                            Campos_validar_gerenciar("Número nota fiscal", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 31: // AE - Chave nota fiscal
                            Campos_validar_gerenciar("Chave nota fiscal", row[column].ToString(), rows, columns, TipoCampoType.Character, 50, false);
                            break;

                        case 32: // AF - Valor de outras despesas
                            Campos_validar_gerenciar("Valor de outras despesas", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 33: // AG - Valor frete
                            Campos_validar_gerenciar("Valor frete", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 34: // AH - Valor desconto
                            Campos_validar_gerenciar("Valor desconto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 35: // AI - Valor impostos adicionais
                            Campos_validar_gerenciar("Valor impostos adicionais", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 36: // AJ - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;

                        case 37: // AK - Código veículo Solution
                            Campos_validar_gerenciar("Código veículo Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 38: // AL - Código veículo legado
                            Campos_validar_gerenciar("Código veículo legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 39: // AM - Número serie veículo
                            Campos_validar_gerenciar("Número serie veículo", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 40: // AN - Classificação
                            Campos_validar_gerenciar("Classificação", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 41: // AO - Hodometro
                            Campos_validar_gerenciar("Hodometro", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, false);
                            break;

                        case 42: // AP - Horimetro
                            Campos_validar_gerenciar("Horimetro", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, false);
                            break;

                        case 43: // AQ - Mecanico
                            Campos_validar_gerenciar("Mecanico", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 44: // AR - Tipo ordem serviço
                            Campos_validar_gerenciar("Tipo ordem serviço", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 45: // AS - Descrição problema
                            Campos_validar_gerenciar("Descrição problema", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 46: // AT - Opinião do problema
                            Campos_validar_gerenciar("Opinião do problema", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 47: // AU - Solução problema
                            Campos_validar_gerenciar("Solução problema", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 48: // AV - Total km rodados
                            Campos_validar_gerenciar("Total km rodados", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.1, false);
                            break;

                        case 49: // AW - Total valor deslocamento
                            Campos_validar_gerenciar("Total valor deslocamento", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 50: // AX - Total valor KM
                            Campos_validar_gerenciar("Total valor KM", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 51: // AY - Total valor serviços
                            Campos_validar_gerenciar("Total valor serviços", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 52: // AZ - Total valor serviço de terceiros
                            Campos_validar_gerenciar("Total valor serviço de terceiros", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 53: // BA - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;
                    }

                    if (columns > 53)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Legado_pedidos_itens(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código item*
                            Campos_validar_gerenciar("Código item", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 2: // B - Código legado item*
                            Campos_validar_gerenciar("Código legado item", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 3: // C - Código legado pedido*
                            Campos_validar_gerenciar("Código legado pedido", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 4: // D - Código pedido Solution
                            Campos_validar_gerenciar("Código pedido Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 5: // E - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 6: // F - Tipo item*
                            Campos_validar_gerenciar("Tipo item", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_item), true);
                            break;

                        case 7: // G - Código produto Solution
                            Campos_validar_gerenciar("Código produto Solution", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 8: // H - Código produto legado*
                            Campos_validar_gerenciar("Código produto legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 9: // I - Descrição produto
                            Campos_validar_gerenciar("Descrição produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 10: // J - Data hora alocação
                            Campos_validar_gerenciar("Data hora alocação", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 11: // K - Unidade
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, false);
                            break;

                        case 12: // L - Código item pedido fornecedor
                            Campos_validar_gerenciar("Código item pedido fornecedor", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 13: // M - Número pedido fornecedor
                            Campos_validar_gerenciar("Número pedido fornecedor", row[column].ToString(), rows, columns, TipoCampoType.Character, 15, false);
                            break;

                        case 14: // N - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, true);
                            break;

                        case 15: // O - Preço unitário
                            Campos_validar_gerenciar("Preço unitário", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 16: // P - Valor desconto
                            Campos_validar_gerenciar("Valor desconto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 17: // Q - Valor frete
                            Campos_validar_gerenciar("Valor frete", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 18: // R - Valor impostos adicionais
                            Campos_validar_gerenciar("Valor impostos adicionais", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 19: // S - Valor outras despesas
                            Campos_validar_gerenciar("Valor outras despesas", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 20: // T - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;

                        case 21: // U - Tipo calculo
                            Campos_validar_gerenciar("Tipo calculo", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 22: // V - Total horas trabalhadas
                            Campos_validar_gerenciar("Total horas trabalhadas", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.8, false);
                            break;

                        case 23: // W - Total horas vendidas
                            Campos_validar_gerenciar("Total horas vendidas", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.8, false);
                            break;

                        case 24: // X - Mecanico
                            Campos_validar_gerenciar("Mecanico", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 25: // Y - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;
                    }

                    if (columns > 25)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Legado_movimentacao(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código empresa Solution*
                            Campos_validar_gerenciar("Código empresa Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, true);
                            break;

                        case 2: // B - Código filial Solution*
                            Campos_validar_gerenciar("Código filial Solution", row[column].ToString(), rows, columns, TipoCampoType.Integer, 2, true);
                            break;

                        case 3: // C - CNPJ Filial
                            Campos_validar_gerenciar("CNPJ Filial", row[column].ToString(), rows, columns, TipoCampoType.Character, 18, false);
                            break;

                        case 4: // D - Código produto Solution
                            Campos_validar_gerenciar("Código produto Solution", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 5: // E - Código produto legado*
                            Campos_validar_gerenciar("Código produto legado", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, true);
                            break;

                        case 6: // F - Grupo/classificação produto
                            Campos_validar_gerenciar("Grupo/classificação produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 7: // G - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, TipoCampoType.Character, 100, false);
                            break;

                        case 8: // H - Tipo movimentação*
                            Campos_validar_gerenciar("Tipo movimentação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_movimentacao), true);
                            break;

                        case 9: // I - Movimenta estoque*
                            Campos_validar_gerenciar("Movimenta estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controla_estoque), true);
                            break;

                        case 10: // J - Número documento
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 11: // K - Data movimentação
                            Campos_validar_gerenciar("Data movimentação", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 12: // L - hora movimentação
                            Campos_validar_gerenciar("hora movimentação", row[column].ToString(), rows, columns, TipoCampoType.DateFormat, 7, false);
                            break;

                        case 13: // M - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.4, true);
                            break;

                        case 14: // N - Custo médio total
                            Campos_validar_gerenciar("Custo médio total", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 15: // O - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, true);
                            break;
                    }

                    if (columns > 15)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Grupos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Grupo ID*
                            Campos_validar_gerenciar("Grupo ID", row[column].ToString(), rows, columns, TipoCampoType.Nivel, 8, true);
                            break;

                        case 2: // C - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: // D - Situação*
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Situacao_grupos), false);
                            break;

                        case 4: // E - Área*
                            Campos_validar_gerenciar("Área", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Area), false);
                            break;

                        case 5: // F - Coeficiente mínimo
                            Campos_validar_gerenciar("Coeficiente mínimo", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 7.4, false);
                            break;

                        case 6: // G - ID do centro de custo
                            Campos_validar_gerenciar("ID do centro de custo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, false);
                            break;

                        case 7: // H - Margem de lucro
                            Campos_validar_gerenciar("Margem de lucro", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 8.4, false);
                            break;

                        case 8: // I - Tipo
                            Campos_validar_gerenciar("Tipo", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_grupo), false);
                            break;

                        case 9: // J - Inutilizado
                            break;

                        case 10: // K - Tipo de Calculo do Preço de Venda
                            Campos_validar_gerenciar("Tipo de Calculo do Preço de Venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 11: // L - Tipo de Cálculo do Preço de Venda Sugerido
                            Campos_validar_gerenciar("Tipo de Cálculo do Preço de Venda Sugerido", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 12: // M - Cód. Tributação Padrão
                            Campos_validar_gerenciar("Cód. Tributação Padrão", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 13: // N - Coeficiente Preço de venda
                            Campos_validar_gerenciar("Coeficiente Preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 7.4, false);
                            break;

                        case 14: // P - Tipo da base do preço de venda
                            Campos_validar_gerenciar("Tipo da base do preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Character, 2, false);
                            break;

                        case 15: // Q - Inutilizado
                            break;

                        case 16: // R - Preço Sugerido
                            Campos_validar_gerenciar("Preço Sugerido", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 17: // S - Coeficiente
                            Campos_validar_gerenciar("Coeficiente", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 7.4, false);
                            break;
                    }

                    if (columns > 17)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Sub_grupos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Subgrupo*
                            Campos_validar_gerenciar("Grupo ID", row[column].ToString(), rows, columns, TipoCampoType.Nivel, 8, true);
                            break;

                        case 2: // C - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: // D - Nível*
                            Campos_validar_gerenciar("Nível", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Nivel), true);
                            break;

                        case 4: // E - Situação*
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Situacao_grupos), true);
                            break;
                    }

                    if (columns > 4)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Produtos(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código do Produto*
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, true);
                            break;

                        case 2: //B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: //C - Descrição adicional do item
                            Campos_validar_gerenciar("Descrição adicional do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 4: //D - Tipo de mercadoria(programa de excelência em gestão)
                            Campos_validar_gerenciar("Tipo de mercadoria", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 5: //E - Marca*
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, true);
                            break;

                        case 6: //F - Departamento
                            Campos_validar_gerenciar("Departamento", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 7: //G - Controla estoque
                            Campos_validar_gerenciar("Controla estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controla_estoque), false);
                            break;

                        case 8: //H - Base de cálculo para COFINS
                            Campos_validar_gerenciar("Base de cálculo para COFINS", row[column].ToString(), rows, columns, TipoCampoType.Character, 1, false);
                            break;

                        case 9: //I - Base de cálculo para PIS
                            Campos_validar_gerenciar("Base de cálculo para PIS", row[column].ToString(), rows, columns, TipoCampoType.Character, 1, false);
                            break;

                        case 10: //J - Situação tributária fiscal - ICMS = CST
                            Campos_validar_gerenciar("Base de cálculo para PIS", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 11: //K - Código do grupo*
                            Campos_validar_gerenciar("Código do grupo", row[column].ToString(), rows, columns, TipoCampoType.Integer, 10, true);
                            break;

                        case 12: //L - % IPI na venda
                            Campos_validar_gerenciar("% IPI na venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 13: //M - % IPI na compra (caso se credite)
                            Campos_validar_gerenciar("% IPI na compra", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 14: //N - Produto Perkins
                            Campos_validar_gerenciar("Produto Perkins", row[column].ToString(), rows, columns, TipoCampoType.Character, 1, false);
                            break;

                        case 15: //O - Peso liquido
                            Campos_validar_gerenciar("Peso Liquido", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 16: //P - Peso bruto
                            Campos_validar_gerenciar("Peso bruto", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 17: //Q - Margem ICMS substituição (MVA)
                            Campos_validar_gerenciar("Margem ICMS substituição (MVA)", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 18: //R - % ICMS substituição
                            Campos_validar_gerenciar("% ICMS substituição", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 19: //S - Sufixo
                            Campos_validar_gerenciar("Sufixo", row[column].ToString(), rows, columns, TipoCampoType.Character, 3, false);
                            break;

                        case 20: //T - Unidade*
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, true);
                            break;

                        case 21: //U - Aplicação
                            Campos_validar_gerenciar("Aplicação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 22: //V - Apelido
                            Campos_validar_gerenciar("Apelido", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 23: //W - % do PIS - (somente os diferentes 1,65)
                            Campos_validar_gerenciar("% do PIS", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 24: //X - % do COFINS (somente os diferentes 7,60)
                            Campos_validar_gerenciar("% do COFINS", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 6.2, false);
                            break;

                        case 25: //Y - Produto Importado ou Nacional*
                            Campos_validar_gerenciar("Importado ou Nacional", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Importado_nacional), true);
                            break;

                        case 26: //Z - Peça original ou paralela*
                            Campos_validar_gerenciar("Peça original ou paralela", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Original_paralela), true);
                            break;

                        case 27: //AA - Classificação
                            Campos_validar_gerenciar("Peça original ou paralela", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 28: //AB - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 29: //AC - Preço de reposição
                            Campos_validar_gerenciar("Preço de reposição", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 30: //AD - Código de referência
                            Campos_validar_gerenciar("Código de referência", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 31: //AE - CAMPO INUTILIZADO
                            break;

                        case 32: //AF - Item similar
                            Campos_validar_gerenciar("Item similar", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 33: //AG - Estoque mínimo
                            Campos_validar_gerenciar("Estoque mínimo", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 34: //AH - Situação*
                            Campos_validar_gerenciar("Situação", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Situacao), true);
                            break;

                        case 35: //AI - Produto usado*
                            Campos_validar_gerenciar("Produto usado", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Usado), true);
                            break;

                        case 36: //AJ - NCM*
                            Campos_validar_gerenciar("NCM", row[column].ToString(), rows, columns, TipoCampoType.Character, 10, true);
                            break;

                        case 37: //AK - Modelo
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, TipoCampoType.Character, 12, false);
                            break;

                        case 38: //AL - Data de alteração do preço de venda
                            Campos_validar_gerenciar("Data de alteração do preço de venda", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 39: //AM - Quantidade da embalagem
                            Campos_validar_gerenciar("Quantidade da embalagem", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 9.2, false);
                            break;

                        case 40: //AN - Tipo da embalagem
                            Campos_validar_gerenciar("NCM", row[column].ToString(), rows, columns, TipoCampoType.Character, 2, false);
                            break;

                        case 41: //AO - Código interno
                            Campos_validar_gerenciar("Código interno", row[column].ToString(), rows, columns, TipoCampoType.Integer, 9, false);
                            break;

                        case 42: //AP - Tipo SPED*
                            Campos_validar_gerenciar("Tipo SPED", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_sped), true);
                            break;

                        case 43: //AQ - Campo Inutilizado
                            break;

                        case 44: //AR - Valor da ultima compra
                            Campos_validar_gerenciar("Valor da ultima compra", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 12.2, false);
                            break;

                        case 45: //AS - Data da ultima compra
                            Campos_validar_gerenciar("Data da ultima compra", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 46: //AT - Classificação na curva ABC
                            Campos_validar_gerenciar("Classificação na curva ABC", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Curva_abc), false);
                            break;

                        case 47: //AU - Fornecedor “PRINCIPAL” do item
                            Campos_validar_gerenciar("Fornecedor “PRINCIPAL” do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 6, false);
                            break;

                        case 48: //AV - Código Finame do item
                            Campos_validar_gerenciar("Código Finame do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 49: //AW - Código GTIN do item
                            Campos_validar_gerenciar("Código GTIN do item", row[column].ToString(), rows, columns, TipoCampoType.Character, 14, false);
                            break;

                        case 50: //AX - Código ANP
                            Campos_validar_gerenciar("Código ANP", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 51: //AY - Classe produto
                            Campos_validar_gerenciar("Classe produto", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Classe), false);
                            break;

                        case 52: //AZ - Código base
                            Campos_validar_gerenciar("Código base", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 53: //BA - Número de serie
                            Campos_validar_gerenciar("Número de serie", row[column].ToString(), rows, columns, TipoCampoType.Character, 40, false);
                            break;

                        case 54: //BB - CAMPO INUTILIZADO
                            break;

                        case 55: //BC - % margem de lucro
                            Campos_validar_gerenciar("% margem de lucro", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 8.4, false);
                            break;

                        case 56: //BD - Código Tipo de cálculo de venda
                            Campos_validar_gerenciar("Código Tipo de cálculo de venda", row[column].ToString(), rows, columns, TipoCampoType.Character, 4, false);
                            break;

                        case 57: //BE - Código produto único
                            Campos_validar_gerenciar("Código produto único", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 58: //BF - Preço garantia
                            Campos_validar_gerenciar("Preço garantia", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.3, false);
                            break;

                        case 59: //BG - Código bandeira
                            Campos_validar_gerenciar("Código bandeira", row[column].ToString(), rows, columns, TipoCampoType.Integer, 4, false);
                            break;

                        case 60: //BH - Código antigo produto*
                            Campos_validar_gerenciar("Código antigo produto", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 61: //BI - Código produto substituto
                            Campos_validar_gerenciar("Código bandeira", row[column].ToString(), rows, columns, TipoCampoType.Character, 20, false);
                            break;

                        case 62: //BJ - Preço de lista
                            Campos_validar_gerenciar("Preço de lista", row[column].ToString(), rows, columns, TipoCampoType.Numeric, 16.2, false);
                            break;

                        case 63: //BK - Data de cadastro produto
                            Campos_validar_gerenciar("Data de cadastro produto", row[column].ToString(), rows, columns, TipoCampoType.Date, 0, false);
                            break;

                        case 64: //BL - Código Fiscal
                            Campos_validar_gerenciar("Código Fiscal", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, false);
                            break;

                        case 65: //BM - Código de barras
                            Campos_validar_gerenciar("Código de barras", row[column].ToString(), rows, columns, TipoCampoType.Character, 50, false);
                            break;

                        case 66: //BN - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, TipoCampoType.Character, 1200, false);
                            break;

                        case 67: //BO - Controle de estoque
                            Campos_validar_gerenciar("Controle de estoque", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Controle_estoque), false);
                            break;
                    }

                    if (columns > 67)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }
                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Plano_de_Contas(DataTable dataTable, int rows) 
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Conta*
                            Campos_validar_gerenciar("Conta", row[column].ToString(), rows, columns, TipoCampoType.Integer, 6, true);
                            break;

                        case 2: // C - Nome*
                            Campos_validar_gerenciar("Nome", row[column].ToString(), rows, columns, TipoCampoType.Character, 60, true);
                            break;

                        case 3: // D - Sintética ou Analítica*
                            Campos_validar_gerenciar("Sintética ou Analítica", row[column].ToString(), rows, columns, TipoCampoType.Dominio, Dominio_retornar(DominioType.Sintetica_analitica), true);
                            break;

                        case 4: // E - Sintética*
                            Campos_validar_gerenciar("Sintética", row[column].ToString(), rows, columns, TipoCampoType.Sintetica, 0, true);
                            break;

                        case 5: // E - Conta Referencial
                            Campos_validar_gerenciar("Conta Referencial", row[column].ToString(), rows, columns, TipoCampoType.Character, 19, false);
                            break;
                    }

                    if (columns > 5)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }

        public void Contas(DataTable dataTable, int rows)
        {
            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    var camposConfiguracao = new Dictionary<int, (string Nome, TipoCampoType Tipo, double Tamanho, bool Obrigatorio)>
                    {
                        {1, ("Código", TipoCampoType.Integer, 6, true)},
                        {2, ("Nome", TipoCampoType.Character, 60, true)},
                        {3, ("Endereço (rua/logradouro)", TipoCampoType.Character, 60, false)},
                        {4, ("Número", TipoCampoType.Integer, 6, false)},
                        {5, ("Cidade", TipoCampoType.Character, 50, false)},
                        {6, ("Bairro", TipoCampoType.Character, 30, false)},
                        {7, ("Complemento", TipoCampoType.Character, 60, false)},
                        {8, ("UF", TipoCampoType.Character, 2, false)},
                        {9, ("CEP", TipoCampoType.Character, 9, false)},
                        {10, ("Código do município", TipoCampoType.Integer, 7, false)},
                        {11, ("Vendedor", TipoCampoType.Integer, 3, false)},
                        {12, ("Pessoa Física/Jurídica", TipoCampoType.Dominio, Dominio_retornar(DominioType.Fisica_juridica), true)},
                        {13, ("CNPJ/CPF", TipoCampoType.Character, 18, true)},
                        {14, ("Inscrição Estadual", TipoCampoType.Character, 20, false)},
                        {15, ("Inscrição Municipal", TipoCampoType.Character, 20, false)},
                        {16, ("Inscrição Suframa", TipoCampoType.Character, 20, false)},
                        {17, ("Classificação", TipoCampoType.Character, 20, false)},
                        {18, ("Segmento", TipoCampoType.Character, 20, false)},
                        {19, ("Data nascimento/Data fundação", TipoCampoType.DateFormat, 4, false)},
                        {20, ("Limite de Credito", TipoCampoType.Numeric, 14.2, false)},
                        {21, ("Situação Ativa / Inativa", TipoCampoType.Dominio, Dominio_retornar(DominioType.Contas_situacao), false)},
                        {22, ("Tipo - Cliente / Fornecedor", TipoCampoType.Dominio, Dominio_retornar(DominioType.Cliente_fornecedor), true)},
                        {23, ("Site", TipoCampoType.Character,60, false)},
                        {24, ("Política de preços", TipoCampoType.Character, 3, false)},
                        {25, ("Sexo", TipoCampoType.Character, 1, false)},
                        {26, ("E-mail", TipoCampoType.Character, 500, false)},
                        {27, ("Contribuinte", TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_contribuinte), false)},
                        {28, ("Regime tributário federal", TipoCampoType.Dominio, Dominio_retornar(DominioType.Regime_tributario), false)},
                        //{29, ("CAMPO INUTILIZADO 1", TipoCampoType.Character, 15, false)},
                        {30, ("Produtor Rural", TipoCampoType.Dominio, Dominio_retornar(DominioType.Sim_nao), false)},
                        {31, ("Número da Inscrição Rural", TipoCampoType.Character, 20, false)},
                        //{32, ("CAMPO INUTILIZADO 2", TipoCampoType.Character, 50, false)},
                        {33, ("Titulo de Eleitor", TipoCampoType.Character, 20, false)},
                        {34, ("Nome fantasia", TipoCampoType.Character, 60, false)},
                        {35, ("Nome do contato", TipoCampoType.Character, 60, false)},
                        {36, ("Data de inclusão", TipoCampoType.DateFormat, 4, false)},
                        {37, ("RG", TipoCampoType.Character, 20, false)},
                        {38, ("Cód. do Representante", TipoCampoType.Integer, 3, false)},
                        {39, ("Cód. conta no Escritório Contábil", TipoCampoType.Character, 20, false)},
                        {40, ("Cód. Da Região", TipoCampoType.Integer, 2, false)},
                        {41, ("Observação", TipoCampoType.Character, 1200, false)},
                        {42, ("Telefone 1", TipoCampoType.Character, 14, false)},
                        {43, ("Telefone 2", TipoCampoType.Character, 14, false)},
                        {44, ("Celular", TipoCampoType.Character, 14, false)},
                        {45, ("Fax", TipoCampoType.Character, 14, false)},
                        {46, ("Código Febraban", TipoCampoType.Character, 3, false)},
                        //{47, ("CAMPO INUTILIZADO 3", TipoCampoType.Integer, 6, false)},
                        {48, ("Agência", TipoCampoType.Character, 10, false)},
                        {49, ("Dígito da Agência", TipoCampoType.Character, 2, false)},
                        {50, ("Conta Corrente", TipoCampoType.Character, 20, false)},
                        {51, ("Dígito Conta Corrente", TipoCampoType.Character, 2, false)},
                        {52, ("Observações Pagamento", TipoCampoType.Character, 1200, false)},
                        {53, ("Filiação Mãe", TipoCampoType.Character, 60, false)},
                        {54, ("Filiação Pai", TipoCampoType.Character, 60, false)},
                        {55, ("RG Expedidor", TipoCampoType.Character, 20, false)},
                        {56, ("RG Data de expedição", TipoCampoType.DateFormat, 4, false)},
                        {57, ("RG UF de expedição", TipoCampoType.Character, 2, false)},
                        {58, ("E-mail NF-e", TipoCampoType.Character, 500, false)},
                        {59, ("Estado Civil", TipoCampoType.Dominio, Dominio_retornar(DominioType.Estado_civil), false)},
                        {60, ("Código Externo HPRO", TipoCampoType.Character, 20, false)},
                        {61, ("Nacionalidade", TipoCampoType.Character, 20, false)},
                        {62, ("Naturalidade", TipoCampoType.Character, 40, false)},
                        {63, ("Validade do Cadastro", TipoCampoType.DateFormat, 4, false)},
                        {64, ("Tipo de Fornecedor", TipoCampoType.Dominio, Dominio_retornar(DominioType.Tipo_fornecedor), false)},
                        {65, ("Jurídico", TipoCampoType.Dominio, Dominio_retornar(DominioType.Sim_nao), false)},
                        {66, ("Filial do cooperado", TipoCampoType.Integer, 2, false)},
                        {67, ("Celular 2", TipoCampoType.Character, 15, false)},
                        {68, ("Telefone 1 ramal", TipoCampoType.Character, 4, false)},
                        {69, ("Telefone 2 ramal", TipoCampoType.Character, 4, false)},
                        {70, ("E-mail Boleto", TipoCampoType.Character, 500, true)},
                        {71, ("Indicador IE", TipoCampoType.Dominio, Dominio_retornar(DominioType.Indicador_ie), false)},
                        {72, ("Politica de Prazo", TipoCampoType.Integer, 4, false)},
                        {73, ("Aviso", TipoCampoType.Character, 1200, false)}
                    };

                    if (camposConfiguracao.TryGetValue(columns, out var config))
                    {
                        Campos_validar_gerenciar(config.Nome, row[column].ToString(), rows, columns, config.Tipo, config.Tamanho, config.Obrigatorio);
                    }

                    if (columns > 73)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }
        }
    }
}