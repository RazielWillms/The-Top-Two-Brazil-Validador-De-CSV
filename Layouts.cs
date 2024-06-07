using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        public void Maquinas(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código do Produto*
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 2: //B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 3: //C - Descrição adicional do item*
                            Campos_validar_gerenciar("Descrição adicional do item", row[column].ToString(), rows, columns, "char", 1200, true);
                            break;

                        case 4: //D - Tipo de mercadoria(programa de excelência em gestão)
                            Campos_validar_gerenciar("Tipo de mercadoria", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 5: //E - Marca
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 6: //F - Departamento
                            Campos_validar_gerenciar("Departamento", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 7: //G - Controla estoque
                            List<String> controla_estoque = new List<String> { "S", "N" };
                            Dominio_validar("Controla estoque", row[column].ToString(), rows, columns, controla_estoque, false);
                            break;

                        case 8: //H - Código do grupo*
                            Campos_validar_gerenciar("Departamento", row[column].ToString(), rows, columns, "integer", 10, true);
                            break;

                        case 9: //I - Peso liquido
                            Campos_validar_gerenciar("Pedo Liquido", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 10: //J - Peso bruto
                            Campos_validar_gerenciar("Peso bruto", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 11: //K - Unidade*
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, "char", 2, true);
                            break;

                        case 12: //L - Aplicação
                            Campos_validar_gerenciar("Aplicação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 13: //M - Apelido
                            Campos_validar_gerenciar("Apelido", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 14: //N - Produto Importado ou Nacional
                            List<String> dom_importado_nacional = new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8" };
                            Dominio_validar("Importado ou Nacional", row[column].ToString(), rows, columns, dom_importado_nacional, false);
                            break;

                        case 15: //O - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 16: //P - Preço de reposição
                            Campos_validar_gerenciar("Preço de reposição", row[column].ToString(), rows, columns, "numeric", 12.2, false);
                            break;

                        case 17: //Q - Código de referência
                            Campos_validar_gerenciar("Código de referência", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 18: //R - Situação
                            List<String> dom_situacao = new List<String> { "A", "I" };
                            Dominio_validar("situacao", row[column].ToString(), rows, columns, dom_situacao, false);
                            break;

                        case 19: //S - Produto usado*
                            List<String> dom_usado = new List<String> { "1", "0" };
                            Dominio_validar("Produto usado", row[column].ToString(), rows, columns, dom_usado, true);
                            break;

                        case 20: //T - NCM*
                            Campos_validar_gerenciar("NCM", row[column].ToString(), rows, columns, "char", 10, true);
                            break;

                        case 21: //U - Modelo
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 22: //V - Classe produto*
                            List<String> dom_classe = new List<String> { "N", "B" };
                            Dominio_validar("Classe", row[column].ToString(), rows, columns, dom_classe, true);
                            break;

                        case 23: //W - Código base*
                            Campos_validar_gerenciar("Código base", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 24: //X - Número de serie
                            Campos_validar_gerenciar("Número de serie", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 25: //Y - Código antigo produto*
                            Campos_validar_gerenciar("Código antigo produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 26: //Z - Código Fiscal
                            Campos_validar_gerenciar("Código Fiscal", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 27: //AB - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 28: //AC - Controle de estoque*
                            List<String> dom_controle = new List<String> { "I" };
                            Dominio_validar("Controle de estoque", row[column].ToString(), rows, columns, dom_controle, true);
                            break;

                        case 29: //AD - Campo Livre
                            Campos_validar_gerenciar("Campo Livre", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 30: //AE - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 31: //AF - Código bandeira*
                            Campos_validar_gerenciar("Código bandeira", row[column].ToString(), rows, columns, "integer", 9, true);
                            break;
                    }

                    if (columns > 31)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }
                Progresso_atualizar(total, rows);

                rows++;
            }

            Progresso_gerenciar(false);
        }

        public void Saldos_maquinas(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Código do Produto*
                            Campos_validar_gerenciar("Código do Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 12.4, true);
                            break;

                        case 4: //D - Valor do Estoque*
                            Campos_validar_gerenciar("Valor do Estoque", row[column].ToString(), rows, columns, "numeric", 12.2, true);
                            break;

                        case 5: //E - Código da prateleira
                            Campos_validar_gerenciar("Código da prateleira", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 6: //F - Data da última compra
                            Campos_validar_gerenciar("Data da última compra", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 7: //G - Valor da última compra
                            Campos_validar_gerenciar("Valor da última compra", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 8: //H - Estoque mínimo
                            Campos_validar_gerenciar("Estoque mínimo", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 9: //I - Descrição
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 10: //J - Código produto único
                            Campos_validar_gerenciar("Código produto único", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 11: //K - Custo Reposição
                            Campos_validar_gerenciar("Estoque mínimo", row[column].ToString(), rows, columns, "numeric", 15.2, false);
                            break;

                        case 12: //L - Preço de venda
                            Campos_validar_gerenciar("Preço de venda", row[column].ToString(), rows, columns, "numeric", 15.3, false);
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

            Progresso_gerenciar(false);
        }

        public void Adiantamentos(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Conta legado*
                            Campos_validar_gerenciar("Conta legado", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Valor do adiantamento*
                            Campos_validar_gerenciar("Valor do adiantamento", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 4: //D - Tipo do adiantamento*
                            List<String> dom_tipo_adiantamento = new List<String> { "C", "F" };
                            Dominio_validar("Tipo do adiantamento", row[column].ToString(), rows, columns, dom_tipo_adiantamento, true);
                            break;

                        case 5: //E - Centro de Custo
                            Campos_validar_gerenciar("Centro de Custo", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 6: //F - Número
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: //G - Observação
                            Campos_validar_gerenciar("Conta legado", row[column].ToString(), rows, columns, "char", 1200, false);
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

            Progresso_gerenciar(false);
        }

        public void Orcamento_balcao(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código Pedido*
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, "integer", 9, true);
                            break;

                        case 2: //B - Código do cliente (sistema antigo)*
                            Campos_validar_gerenciar("Código Legado do Cliente", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 3: //C - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 4: //D - Política Prazo
                            Campos_validar_gerenciar("Política Prazo", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 5: //E - Politica Preço
                            Campos_validar_gerenciar("Politica Preço", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 6: //F - Tipo Operação
                            List<String> dom_tipo_operacao = new List<String> { "V", "S", "E", "C", "D" };
                            Dominio_validar("Tipo Operação", row[column].ToString(), rows, columns, dom_tipo_operacao, false);
                            break;

                        case 7: //G - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 8: //H - Funcionário Abertura O.C
                            Campos_validar_gerenciar("Funcionário Abertura O.C", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 9: //I - Data Validade
                            Campos_validar_gerenciar("Data Validade", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 10: //J - Data Abertura*
                            Campos_validar_gerenciar("Data Abertura", row[column].ToString(), rows, columns, "date", 0, true);
                            break;

                        case 11: //K - Data Parcelamento
                            Campos_validar_gerenciar("Data Parcelamento", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 12: //L - Situação*
                            List<String> dom_orc_situacao = new List<String> { "A", "F" };
                            Dominio_validar("Situação", row[column].ToString(), rows, columns, dom_orc_situacao, true);
                            break;

                        case 13: //M - Status*
                            List<String> dom_status = new List<String> { "A", "P", "C", "F", "B", "S", "X", "Y" };
                            Dominio_validar("Status", row[column].ToString(), rows, columns, dom_status, true);
                            break;

                        case 14: //N - Produto*
                            Campos_validar_gerenciar("Produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 15: //O - Descrição Produto
                            Campos_validar_gerenciar("Descrição Produto", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 16: //P - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 17: //Q - Preço Unitário*
                            Campos_validar_gerenciar("Preço Unitário", row[column].ToString(), rows, columns, "numeric", 16.3, true);
                            break;

                        case 18: //R - Valor Desconto
                            Campos_validar_gerenciar("Valor Desconto", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 19: //S - Vendedor Produto
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, "integer", 6, false);
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

            Progresso_gerenciar(false);
        }

        public void Orcamento_oficina(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Número*
                            Campos_validar_gerenciar("Número", row[column].ToString(), rows, columns, "integer", 9, true);
                            break;

                        case 2: //B - Código da Filial Solution*
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 3: //C - ID do Veículo*
                            Campos_validar_gerenciar("ID do Veículo", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 4: //D - Série do veículo*
                            Campos_validar_gerenciar("Série do veículo", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 5: //E - Conta do cliente legado - sistema antigo*
                            Campos_validar_gerenciar("Conta do cliente legado", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 6: //F - Tipo da OS
                            Campos_validar_gerenciar("Tipo da OS", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: //G - Data de abertura
                            Campos_validar_gerenciar("Data de abertura", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 8: //H - ID do mecânico no Solution
                            Campos_validar_gerenciar("Mecânico no Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 9: //I - ID do vendedor no Solution
                            Campos_validar_gerenciar("Vendedor no Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 10: //J - ID do local de venda
                            Campos_validar_gerenciar("local de venda", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 11: //K - ID da política de preço
                            Campos_validar_gerenciar("política de preço", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 12: //L - ID da política de prazo
                            Campos_validar_gerenciar("política de prazo", row[column].ToString(), rows, columns, "char", 3, true);
                            break;

                        case 13: //M - Código do produto*
                            Campos_validar_gerenciar("Código do produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 14: //N - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 15: //O - Preço unitário*
                            Campos_validar_gerenciar("Preço unitário", row[column].ToString(), rows, columns, "numeric", 16.2, false);
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

            Progresso_gerenciar(false);
        }

        public void Estatisticas(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: //A - Código filial Solution*
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: //B - Código produto*
                            Campos_validar_gerenciar("Código produto", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 3: //C - Data movimetação (mês e ano)*
                            Campos_validar_gerenciar("Data movimetação", row[column].ToString(), rows, columns, "date", 0, true);
                            break;

                        case 4: //D - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 15.4, true);
                            break;

                        case 5: //E - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
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

            Progresso_gerenciar(false);
        }

        public void Veiculos_clientes(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código*
                            Campos_validar_gerenciar("Código", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 2: // B - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 3: // C - Placa
                            Campos_validar_gerenciar("Placa", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 4: // D - Meses Garantia
                            Campos_validar_gerenciar("Meses Garantia", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 5: // E - Hrs.Garantia
                            Campos_validar_gerenciar("Hrs.Garantia", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 6: // F - Km garantia
                            Campos_validar_gerenciar("Km garantia", row[column].ToString(), rows, columns, "numeric", 10.1, false);
                            break;

                        case 7: // G - Novo Usado*
                            List<String> dom_novo_usado = new List<String> { "N", "U" };
                            Dominio_validar("Novo Usado", row[column].ToString(), rows, columns, dom_novo_usado, true);
                            break;

                        case 8: // H - Versão
                            Campos_validar_gerenciar("Versão", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 9: // I - Ano fabricação*
                            Campos_validar_gerenciar("Ano fabricação", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 10: // J - Ano modelo*
                            Campos_validar_gerenciar("Ano modelo", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 11: // K - Código da conta de cliente (sistema antigo)*
                            Campos_validar_gerenciar("Código da conta de cliente (sistema antigo)", row[column].ToString(), rows, columns, "char", 6, true);
                            break;

                        case 12: // L - Modelo*
                            Campos_validar_gerenciar("Modelo", row[column].ToString(), rows, columns, "char", 12, true);
                            break;

                        case 13: // M - numero NF de compra
                            Campos_validar_gerenciar("numero NF de compra", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 14: // N - Data de compra
                            Campos_validar_gerenciar("Data de compra", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 15: // O - Código da conta de fornecedor
                            Campos_validar_gerenciar("Código da conta de fornecedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 16: // P - Nome fornecedor
                            Campos_validar_gerenciar("Nome fornecedor", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 17: // Q - Código produto estoque
                            Campos_validar_gerenciar("Código produto estoque", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 18: // R - Numero de serie*
                            Campos_validar_gerenciar("Numero de serie", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 19: // S - Serie motor*
                            Campos_validar_gerenciar("Serie motor", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 20: // T - Série da bomba hidráulica
                            Campos_validar_gerenciar("Série da bomba hidráulica", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 21: // U - Série de transmissão
                            Campos_validar_gerenciar("Série de transmissão", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 22: // V - Série da caixa de câmbio
                            Campos_validar_gerenciar("Série da caixa de câmbio", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 23: // W - Série da bomba injetora
                            Campos_validar_gerenciar("Série da bomba injetora", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 24: // X - Série do monobloco
                            Campos_validar_gerenciar("Série do monobloco", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Série do eixo dianteiro
                            Campos_validar_gerenciar("Série do eixo dianteiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 26: // Z - Série da plataforma
                            Campos_validar_gerenciar("Série da plataforma", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 27: // AA - Pneus dianteiro
                            Campos_validar_gerenciar("Pneus dianteiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 28: // AB - Pneus traseiro
                            Campos_validar_gerenciar("Pneus traseiro", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 29: // AC - Série direção hidráulica
                            Campos_validar_gerenciar("Série direção hidráulica", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 30: // AD - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 200, false);
                            break;

                        case 31: // AE - Tipo equipamento*
                            List<String> dom_tipo_equipamento = new List<String> { "#", "J", "8", "4", "A", "5", "N", "C", "R", "D", "2", "L", "K", "P", "H", "V", "I", "3", "S", "6", "M", "O", "9", "Z", "B", "U", "F", "7", "Y", "T", "G", "Q", "1", "E", "X" };
                            Dominio_validar("Tipo equipamento", row[column].ToString(), rows, columns, dom_tipo_equipamento, true);
                            break;

                        case 32: // AF - Código do pedido da gestão de compra
                            Campos_validar_gerenciar("Código do pedido da gestão de compra", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 33: // AG - Cor código*
                            Campos_validar_gerenciar("Cor código", row[column].ToString(), rows, columns, "char", 4, true);
                            break;

                        case 34: // AH - Cor descrição*
                            Campos_validar_gerenciar("Cor descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 35: // AI - Potência do Motor (CV)
                            Campos_validar_gerenciar("Potência do Motor (CV)", row[column].ToString(), rows, columns, "numeric", 8.1, false);
                            break;

                        case 36: // AJ - CM3 (cilindradas)
                            Campos_validar_gerenciar("CM3 (cilindradas)", row[column].ToString(), rows, columns, "numeric", 8.1, false);
                            break;

                        case 37: // AK - Peso líquido (KG)
                            Campos_validar_gerenciar("Peso líquido (KG)", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 38: // AL - Peso bruto (KG)
                            Campos_validar_gerenciar("Peso bruto (KG)", row[column].ToString(), rows, columns, "numeric", 10, false);
                            break;

                        case 39: // AM - Tipo combustivel*
                            Campos_validar_gerenciar("Tipo combustivel", row[column].ToString(), rows, columns, "char", 10, true);
                            break;

                        case 40: // AN - CMKG
                            Campos_validar_gerenciar("CMKG", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 41: // AO - TMA
                            Campos_validar_gerenciar("TMA", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 42: // AP - Distância entre eixos (mm)
                            Campos_validar_gerenciar("Distância entre eixos (mm)", row[column].ToString(), rows, columns, "numeric", 8.2, false);
                            break;

                        case 43: // AQ - RENAVAM
                            Campos_validar_gerenciar("RENAVAM", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 44: // AR - Tipo pintura*
                            Campos_validar_gerenciar("Tipo pintura", row[column].ToString(), rows, columns, "char", 1, true);
                            break;

                        case 45: // AS - Tipo de Veículo Renavam/Denatran
                            List<String> dom_tipo_renavam_denatram = new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26" };
                            Dominio_validar("Tipo de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, dom_tipo_renavam_denatram, false);
                            break;

                        case 46: // AT - Espécie de Veículo Renavam/Denatran
                            List<String> dom_especie_veiculo_renavam_denatram = new List<String> { "0", "1", "2", "3", "4", "5", "6" };
                            Dominio_validar("Espécie de Veículo Renavam/Denatran", row[column].ToString(), rows, columns, dom_especie_veiculo_renavam_denatram, false);
                            break;

                        case 47: // AU - Marca Modelo Renavam/Denatran
                            Campos_validar_gerenciar("Marca Modelo Renavam/Denatran", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 48: // AV - Codigo do DN
                            Campos_validar_gerenciar("Codigo do DN", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 49: // AW - Chassis*
                            Campos_validar_gerenciar("Chassis", row[column].ToString(), rows, columns, "char", 100, true);
                            break;

                        case 50: // AX - Marca
                            Campos_validar_gerenciar("Marca", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 51: // AY - Data entrega tecnica
                            Campos_validar_gerenciar("Data entrega tecnica", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 52: // AZ - Data ultima revisão
                            Campos_validar_gerenciar("Data ultima revisão", row[column].ToString(), rows, columns, "date", 0, false);
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

            Progresso_gerenciar(false);
        }

        public void Imobilizado_itens(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código da Empresa Solution*
                            Campos_validar_gerenciar("Código da Empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código da Filial Solution*
                            Campos_validar_gerenciar("Código da Filial Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 3: // C - Código do Item*
                            Campos_validar_gerenciar("Código do Item", row[column].ToString(), rows, columns, "numeric", 6.2, true);
                            break;

                        case 4: // D - Código da Conta (Plano de Contas)
                            Campos_validar_gerenciar("Código da Conta (Plano de Contas)", row[column].ToString(), rows, columns, "char", 11, false);
                            break;

                        case 5: // E - Data do lancto*
                            Campos_validar_gerenciar("Data do lancto", row[column].ToString(), rows, columns, "date", 10, true);
                            break;

                        case 6: // F - Data da aquisição*
                            Campos_validar_gerenciar("Data da aquisição", row[column].ToString(), rows, columns, "date", 10, true);
                            break;

                        case 7: // G - Centro de Custo
                            Campos_validar_gerenciar("Centro de Custo", row[column].ToString(), rows, columns, "char", 6, false);
                            break;

                        case 8: // H - % de Depreciação do Item
                            Campos_validar_gerenciar("% de Depreciação do Item", row[column].ToString(), rows, columns, "numeric", 5.2, false);
                            break;

                        case 9: // I - % de Depreciação Gerencial
                            Campos_validar_gerenciar("% de Depreciação Gerencial", row[column].ToString(), rows, columns, "numeric", 6.2, false);
                            break;

                        case 10: // J - % residual
                            Campos_validar_gerenciar("% residual", row[column].ToString(), rows, columns, "numeric", 5.2, false);
                            break;

                        case 11: // K - Débito ou Crédito*
                            List<String> dom_debito_credito = new List<String> { "D", "C" };
                            Dominio_validar("Débito ou Crédito", row[column].ToString(), rows, columns, dom_debito_credito, true);
                            break;

                        case 12: // L - Chave*
                            List<String> dom_chave = new List<String> { "G", "C" };
                            Dominio_validar("Chave", row[column].ToString(), rows, columns, dom_chave, true);
                            break;

                        case 13: // M - Tipo lançamento
                            List<String> dom_tipo_lanacamento = new List<String> { "A", "T", "I" };
                            Dominio_validar("Tipo lançamento", row[column].ToString(), rows, columns, dom_tipo_lanacamento, false);
                            break;

                        case 14: // N - Tipo Baixa
                            List<String> dom_tipo_baixa = new List<String> { "B", "T" };
                            Dominio_validar("Tipo Baixa", row[column].ToString(), rows, columns, dom_tipo_baixa, false);
                            break;

                        case 15: // O - Número do documento de aquisição
                            Campos_validar_gerenciar("Número do documento de aquisição", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 16: // P - Nome do Fornecedor
                            Campos_validar_gerenciar("Nome do Fornecedor", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 17: // Q - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 225, true);
                            break;

                        case 18: // R - Descrição sucienta da função do bem na atividade do estabelecimento (obrigatório para Sped Fiscal)*
                            Campos_validar_gerenciar("Descrição sucienta", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 19: // S - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 20: // T - Número da Apólice
                            Campos_validar_gerenciar("Número da Apólice", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 21: // U - Data do Vencimento
                            Campos_validar_gerenciar("Data do Vencimento", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 22: // V - Código Externo
                            Campos_validar_gerenciar("Código Externo", row[column].ToString(), rows, columns, "char", 12, false);
                            break;

                        case 23: // W - Código do Local
                            Campos_validar_gerenciar("Código do Local", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 24: // X - Código do Responsável
                            Campos_validar_gerenciar("Código do Responsável", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 25: // Y - Código do tipo do bem
                            Campos_validar_gerenciar("Código do tipo do bem", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 26: // Z - Código da Seguradora
                            Campos_validar_gerenciar("Código da Seguradora", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 27: // AA - Tipo Documento de aquisição
                            Campos_validar_gerenciar("Tipo Documento de aquisição", row[column].ToString(), rows, columns, "integer", 3, false);
                            break;

                        case 28: // AB - Situação do Bem
                            Campos_validar_gerenciar("Situação do Bem", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 29: // AC - Chassis
                            Campos_validar_gerenciar("Chassis", row[column].ToString(), rows, columns, "char", 10, false);
                            break;

                        case 30: // AD - Placa
                            Campos_validar_gerenciar("Placa", row[column].ToString(), rows, columns, "char", 9, false);
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

            Progresso_gerenciar(false);
        }

        public void Imobilizado_saldos(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código da Empresa*
                            Campos_validar_gerenciar("Código da Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código do Item*
                            Campos_validar_gerenciar("Código do Item", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 3: // C - Valor Original*
                            Campos_validar_gerenciar("Valor Original", row[column].ToString(), rows, columns, "numeric", 15.2, true);
                            break;

                        case 4: // D - Valor Original Corrigido
                            Campos_validar_gerenciar("Valor Original Corrigido", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 5: // E - Depreciação Acumulada Corrigido
                            Campos_validar_gerenciar("Depreciação Acumulada Corrigido", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 6: // F - Valor Original Moeda
                            Campos_validar_gerenciar("Valor Original Moeda", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 7: // G - Depreciação acumulada Moeda
                            Campos_validar_gerenciar("Depreciação acumulada Moeda", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 8: // H - Valor Original Ufir
                            Campos_validar_gerenciar("Valor Original Ufir", row[column].ToString(), rows, columns, "numeric", 16.4, false);
                            break;

                        case 9: // I - Depreciação acumulada Ufir
                            Campos_validar_gerenciar("Depreciação acumulada Ufir", row[column].ToString(), rows, columns, "numeric", 16.4, false);
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

            Progresso_gerenciar(false);
        }

        public void Legado_financeiro(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código legado documento*
                            Campos_validar_gerenciar("Código legado documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Número documento*
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Código da conta Solution
                            Campos_validar_gerenciar("Código da conta Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 4: // D - Código da conta legado*
                            Campos_validar_gerenciar("Código da conta legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 5: // E - Código endereço legado
                            Campos_validar_gerenciar("Código endereço legado", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 6: // F - Código endereço Solution
                            Campos_validar_gerenciar("Código endereço Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 7: // G - Tipo de documento*
                            List<String> dom_tipo_documento = new List<String> { "#", "C", "T", "A" };
                            Dominio_validar("Tipo de documento", row[column].ToString(), rows, columns, dom_tipo_documento, true);
                            break;

                        case 8: // H - Pagamento ou recebimento*
                            List<String> dom_pagar_receber = new List<String> { "P", "R" };
                            Dominio_validar("Pagamento ou recebimento", row[column].ToString(), rows, columns, dom_pagar_receber, true);
                            break;

                        case 9: // I - Código empresa Solution*
                            Campos_validar_gerenciar("Código empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 10: // J - Código filial Solution*
                            Campos_validar_gerenciar("Código filial Solution", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 11: // K - CNPJ filial
                            Campos_validar_gerenciar("CNPJ filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 12: // L - Data de emissão*
                            Campos_validar_gerenciar("Data de emissão", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 13: // M - Data de vencimento*
                            Campos_validar_gerenciar("Data de vencimento", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
                            break;

                        case 14: // N - Portador
                            Campos_validar_gerenciar("Portador", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 15: // O - Número da parcela
                            Campos_validar_gerenciar("Número da parcela", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 16: // P - Número nota fiscal
                            Campos_validar_gerenciar("Número nota fiscal", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 17: // Q - Centro de custo
                            Campos_validar_gerenciar("Centro de custo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 18: // R - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 19: // S - Valor*
                            Campos_validar_gerenciar("Valor", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 20: // T - Valor de juros
                            Campos_validar_gerenciar("Valor de juros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 21: // U - Valor de desconto
                            Campos_validar_gerenciar("Valor de desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 22: // V - Valor de multa
                            Campos_validar_gerenciar("Valor de multa", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 23: // W - Número febraban banco
                            Campos_validar_gerenciar("Número febraban banco", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 24: // X - Nosso número boleto
                            Campos_validar_gerenciar("Nosso número boleto", row[column].ToString(), rows, columns, "char", 30, false);
                            break;

                        case 25: // Y - Dias de atraso
                            Campos_validar_gerenciar("Dias de atraso", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 26: // Z - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
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

            Progresso_gerenciar(false);
        }

        public void Legado_pagamentos(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código legado pagamento*
                            Campos_validar_gerenciar("Código legado pagamento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado documento*
                            Campos_validar_gerenciar("Código legado documento", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Número documento
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 4: // D - Código documento Solution
                            List<String> dom_codigo_documento = new List<String> { "", "null", "NULL" };
                            Dominio_validar("Código documento Solution", row[column].ToString(), rows, columns, dom_codigo_documento, false);
                            break;

                        case 5: // E - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 6: // F - CNPJ Filial
                            Campos_validar_gerenciar("CNPJ Filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 7: // G - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 8: // H - Valor*
                            Campos_validar_gerenciar("Valor", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 9: // I - Valor juros
                            Campos_validar_gerenciar("Valor juros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 10: // J - Valor multa
                            Campos_validar_gerenciar("Valor multa", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 11: // K - Desconto valor
                            Campos_validar_gerenciar("Desconto valor", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 12: // L - Data pagamento*
                            Campos_validar_gerenciar("Data pagamento", row[column].ToString(), rows, columns, "date_format", 3, true); //tamanho 3 significa "yyyy/MM/dd"
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

            Progresso_gerenciar(false);
        }

        public void Legado_pedidos(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código pedido*
                            Campos_validar_gerenciar("Código pedido", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado pedido*
                            Campos_validar_gerenciar("Código legado pedido", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 4: // D - Filial*
                            Campos_validar_gerenciar("Filial", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 5: // E - CNPJ filial
                            Campos_validar_gerenciar("CNPJ filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 6: // F - Módulo*
                            List<String> dom_modulo = new List<String> { "5", "17" };
                            Dominio_validar("Módulo", row[column].ToString(), rows, columns, dom_modulo, true);
                            break;

                        case 7: // G - Tipo*
                            List<String> dom_tipo = new List<String> { "O", "P" };
                            Dominio_validar("Tipo", row[column].ToString(), rows, columns, dom_tipo, true);
                            break;

                        case 8: // H - Data hora abertura
                            Campos_validar_gerenciar("Data hora abertura", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 9: // I - Data hora validade
                            Campos_validar_gerenciar("Data hora validade", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 10: // J - Data hora encerramento
                            Campos_validar_gerenciar("Data hora encerramento", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 11: // K - Código cliente legado*
                            Campos_validar_gerenciar("Código cliente legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 12: // L - Código legado endereço
                            Campos_validar_gerenciar("Código legado endereço", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 13: // M - Código endereço Solution
                            Campos_validar_gerenciar("Código endereço Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 14: // N - Código cliente Solution
                            Campos_validar_gerenciar("Código cliente Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 15: // O - Nome cliente
                            Campos_validar_gerenciar("Nome cliente", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 16: // P - Logradouro cliente
                            Campos_validar_gerenciar("Logradouro cliente", row[column].ToString(), rows, columns, "char", 500, false);
                            break;

                        case 17: // Q - Cidade cliente
                            Campos_validar_gerenciar("Cidade cliente", row[column].ToString(), rows, columns, "char", 60, false);
                            break;

                        case 18: // R - UF cliente
                            Campos_validar_gerenciar("UF cliente", row[column].ToString(), rows, columns, "char", 2, false);
                            break;

                        case 19: // S - CEP cliente
                            Campos_validar_gerenciar("CEP cliente", row[column].ToString(), rows, columns, "char", 9, false);
                            break;

                        case 20: // T - CNPJ/CPF cliente
                            Campos_validar_gerenciar("CNPJ/CPF cliente", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 21: // U - Inscrição estadual cliente
                            Campos_validar_gerenciar("Inscrição estadual cliente", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 22: // V - Inscrição municipal cliente
                            Campos_validar_gerenciar("Inscrição municipal cliente", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 23: // W - Vendedor
                            Campos_validar_gerenciar("Vendedor", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 24: // X - Politica prazo
                            Campos_validar_gerenciar("Politica prazo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Tipo pagamento*
                            List<String> dom_pagamento = new List<String> { "V", "P" };
                            Dominio_validar("Tipo pagamento", row[column].ToString(), rows, columns, dom_pagamento, true);
                            break;

                        case 26: // Z - Forma pagamento*
                            List<String> dom_forma_pagamento = new List<String> { "A", "2", "4", "5", "0", "1", "6", "3", "F", "9", "8" };
                            Dominio_validar("Forma pagamento", row[column].ToString(), rows, columns, dom_forma_pagamento, true);
                            break;

                        case 27: // AA - Número parcelas
                            Campos_validar_gerenciar("Número parcelas", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 28: // AB - Data hora parcelamento
                            Campos_validar_gerenciar("Data hora parcelamento", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 29: // AC - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 30: // AD - Número nota fiscal
                            Campos_validar_gerenciar("Número nota fiscal", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 31: // AE - Chave nota fiscal
                            Campos_validar_gerenciar("Chave nota fiscal", row[column].ToString(), rows, columns, "char", 50, false);
                            break;

                        case 32: // AF - Valor de outras despesas
                            Campos_validar_gerenciar("Valor de outras despesas", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 33: // AG - Valor frete
                            Campos_validar_gerenciar("Valor frete", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 34: // AH - Valor desconto
                            Campos_validar_gerenciar("Valor desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 35: // AI - Valor impostos adicionais
                            Campos_validar_gerenciar("Valor impostos adicionais", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 36: // AJ - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 37: // AK - Código veículo Solution
                            Campos_validar_gerenciar("Código veículo Solution", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 38: // AL - Código veículo legado
                            Campos_validar_gerenciar("Código veículo legado", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 39: // AM - Número serie veículo
                            Campos_validar_gerenciar("Número serie veículo", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 40: // AN - Classificação
                            Campos_validar_gerenciar("Classificação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 41: // AO - Hodometro
                            Campos_validar_gerenciar("Hodometro", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 42: // AP - Horimetro
                            Campos_validar_gerenciar("Horimetro", row[column].ToString(), rows, columns, "integer", 10, false);
                            break;

                        case 43: // AQ - Mecanico
                            Campos_validar_gerenciar("Mecanico", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 44: // AR - Tipo ordem serviço
                            Campos_validar_gerenciar("Tipo ordem serviço", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 45: // AS - Descrição problema
                            Campos_validar_gerenciar("Descrição problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 46: // AT - Opinião do problema
                            Campos_validar_gerenciar("Opinião do problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 47: // AU - Solução problema
                            Campos_validar_gerenciar("Solução problema", row[column].ToString(), rows, columns, "char", 1200, false);
                            break;

                        case 48: // AV - Total km rodados
                            Campos_validar_gerenciar("Total km rodados", row[column].ToString(), rows, columns, "numeric", 16.1, false);
                            break;

                        case 49: // AW - Total valor deslocamento
                            Campos_validar_gerenciar("Total valor deslocamento", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 50: // AX - Total valor KM
                            Campos_validar_gerenciar("Total valor KM", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 51: // AY - Total valor serviços
                            Campos_validar_gerenciar("Total valor serviços", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 52: // AZ - Total valor serviço de terceiros
                            Campos_validar_gerenciar("Total valor serviço de terceiros", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 53: // BA - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
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

            Progresso_gerenciar(false);
        }

        public void Legado_pedidos_itens(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {

                    switch (columns)
                    {
                        case 1: // A - Código item*
                            Campos_validar_gerenciar("Código item", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 2: // B - Código legado item*
                            Campos_validar_gerenciar("Código legado item", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 3: // C - Código legado pedido*
                            Campos_validar_gerenciar("Código legado pedido", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 4: // D - Código pedido Solution
                            Campos_validar_gerenciar("Código pedido Solution", row[column].ToString(), rows, columns, "integer", 9, false);
                            break;

                        case 5: // E - Empresa*
                            Campos_validar_gerenciar("Empresa", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 6: // F - Tipo item*
                            List<String> dom_tipo_item = new List<String> { "SP", "P", "ST" };
                            Dominio_validar("Tipo item", row[column].ToString(), rows, columns, dom_tipo_item, true);
                            break;

                        case 7: // G - Código produto Solution
                            Campos_validar_gerenciar("Código produto Solution", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 8: // H - Código produto legado*
                            Campos_validar_gerenciar("Código produto legado", row[column].ToString(), rows, columns, "char", 20, true);
                            break;

                        case 9: // I - Descrição produto
                            Campos_validar_gerenciar("Descrição produto", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 10: // J - Data hora alocação
                            Campos_validar_gerenciar("Data hora alocação", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 11: // K - Unidade
                            Campos_validar_gerenciar("Unidade", row[column].ToString(), rows, columns, "char", 6, false);
                            break;

                        case 12: // L - Código item pedido fornecedor
                            Campos_validar_gerenciar("Código item pedido fornecedor", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 13: // M - Número pedido fornecedor
                            Campos_validar_gerenciar("Número pedido fornecedor", row[column].ToString(), rows, columns, "char", 15, false);
                            break;

                        case 14: // N - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 15: // O - Preço unitário
                            Campos_validar_gerenciar("Preço unitário", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 16: // P - Valor desconto
                            Campos_validar_gerenciar("Valor desconto", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 17: // Q - Valor frete
                            Campos_validar_gerenciar("Valor frete", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 18: // R - Valor impostos adicionais
                            Campos_validar_gerenciar("Valor impostos adicionais", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 19: // S - Valor outras despesas
                            Campos_validar_gerenciar("Valor outras despesas", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 20: // T - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
                            break;

                        case 21: // U - Tipo calculo
                            Campos_validar_gerenciar("Tipo calculo", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 22: // V - Total horas trabalhadas
                            Campos_validar_gerenciar("Total horas trabalhadas", row[column].ToString(), rows, columns, "numeric", 16.8, false);
                            break;

                        case 23: // W - Total horas vendidas
                            Campos_validar_gerenciar("Total horas vendidas", row[column].ToString(), rows, columns, "numeric", 16.8, false);
                            break;

                        case 24: // X - Mecanico
                            Campos_validar_gerenciar("Mecanico", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 25: // Y - Observação
                            Campos_validar_gerenciar("Observação", row[column].ToString(), rows, columns, "char", 1200, false);
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

            Progresso_gerenciar(false);
        }

        public void Legado_movimentacao(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Código empresa Solution*
                            Campos_validar_gerenciar("Código empresa Solution", row[column].ToString(), rows, columns, "integer", 4, true);
                            break;

                        case 2: // B - Código filial Solution*
                            Campos_validar_gerenciar("Código filial Solution", row[column].ToString(), rows, columns, "integer", 2, true);
                            break;

                        case 3: // C - CNPJ Filial
                            Campos_validar_gerenciar("CNPJ Filial", row[column].ToString(), rows, columns, "char", 18, false);
                            break;

                        case 4: // D - Código produto Solution
                            Campos_validar_gerenciar("Código produto Solution", row[column].ToString(), rows, columns, "char", 20, false);
                            break;

                        case 5: // E - Código produto legado*
                            Campos_validar_gerenciar("Código produto legado", row[column].ToString(), rows, columns, "char", 40, true);
                            break;

                        case 6: // F - Grupo/classificação produto
                            Campos_validar_gerenciar("Grupo/classificação produto", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 7: // G - Operação
                            Campos_validar_gerenciar("Operação", row[column].ToString(), rows, columns, "char", 100, false);
                            break;

                        case 8: // H - Tipo movimentação*
                            List<String> dom_tipo_movimentacao = new List<String> { "S", "E" };
                            Dominio_validar("Tipo movimentação", row[column].ToString(), rows, columns, dom_tipo_movimentacao, true);
                            break;

                        case 9: // I - Movimenta estoque*
                            List<String> dom_movimenta_estoque = new List<String> { "S", "N" };
                            Dominio_validar("Movimenta estoque", row[column].ToString(), rows, columns, dom_movimenta_estoque, true);
                            break;

                        case 10: // J - Número documento
                            Campos_validar_gerenciar("Número documento", row[column].ToString(), rows, columns, "char", 40, false);
                            break;

                        case 11: // K - Data movimentação
                            Campos_validar_gerenciar("Data movimentação", row[column].ToString(), rows, columns, "date", 0, false);
                            break;

                        case 12: // L - hora movimentação
                            Campos_validar_gerenciar("hora movimentação", row[column].ToString(), rows, columns, "date_format", 7, false);
                            break;

                        case 13: // M - Quantidade*
                            Campos_validar_gerenciar("Quantidade", row[column].ToString(), rows, columns, "numeric", 16.4, true);
                            break;

                        case 14: // N - Custo médio total
                            Campos_validar_gerenciar("Custo médio total", row[column].ToString(), rows, columns, "numeric", 16.2, false);
                            break;

                        case 15: // O - Valor total*
                            Campos_validar_gerenciar("Valor total", row[column].ToString(), rows, columns, "numeric", 16.2, true);
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

            Progresso_gerenciar(false);
        }
        
        public void Grupos(DataTable dataTable, int rows)
        {

            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Grupo ID*
                            Campos_validar_gerenciar("Grupo ID", row[column].ToString(), rows, columns, "nivel", 8, true);
                            break;

                        case 2: // B - Grupo Solution*
                            Campos_validar_gerenciar("Grupo Solution", row[column].ToString(), rows, columns, "nivel", 8, true);
                            break;

                        case 3: // C - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 4: // D - Situação*
                            List<String> dom_situacao = new List<String> { "A" };
                            Dominio_validar("Situação", row[column].ToString(), rows, columns, dom_situacao, false);
                            break;

                        case 5: // E - Área*
                            List<String> dom_area = new List<String> { "1" };
                            Dominio_validar("Área", row[column].ToString(), rows, columns, dom_area, false);
                            break;

                        case 6: // F - Coeficiente mínimo
                            Campos_validar_gerenciar("Coeficiente mínimo", row[column].ToString(), rows, columns, "numeric", 7.4, false);
                            break;

                        case 7: // G - ID do centro de custo
                            Campos_validar_gerenciar("ID do centro de custo", row[column].ToString(), rows, columns, "integer", 6, false);
                            break;

                        case 8: // H - Margem de lucro
                            Campos_validar_gerenciar("Margem de lucro", row[column].ToString(), rows, columns, "numeric", 8.4, false);
                            break;

                        case 9: // I - Tipo
                            List<String> dom_tipo = new List<String> { "E" };
                            Dominio_validar("Tipo", row[column].ToString(), rows, columns, dom_tipo, false);
                            break;

                        case 10: // J - Inutilizado
                            break;

                        case 11: // K - Tipo de Calculo do Preço de Venda
                            Campos_validar_gerenciar("Tipo de Calculo do Preço de Venda", row[column].ToString(), rows, columns, "numeric", 6.2, false);
                            break;

                        case 12: // L - Tipo de Cálculo do Preço de Venda Sugerido
                            Campos_validar_gerenciar("Tipo de Cálculo do Preço de Venda Sugerido", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 13: // M - Cód. Tributação Padrão
                            Campos_validar_gerenciar("Cód. Tributação Padrão", row[column].ToString(), rows, columns, "char", 3, false);
                            break;

                        case 14: // N - Coeficiente Preço de venda
                            Campos_validar_gerenciar("Coeficiente Preço de venda", row[column].ToString(), rows, columns, "numeric", 7.4, false);
                            break;

                        case 15: // P - Tipo da base do preço de venda
                            Campos_validar_gerenciar("Tipo da base do preço de venda", row[column].ToString(), rows, columns, "char", 2, false);
                            break;

                        case 16: // Q - Inutilizado
                            break;

                        case 17: // R - Preço Sugerido
                            Campos_validar_gerenciar("Preço Sugerido", row[column].ToString(), rows, columns, "integer", 4, false);
                            break;

                        case 18: // S - Coeficiente
                            Campos_validar_gerenciar("Coeficiente", row[column].ToString(), rows, columns, "numeric", 7.4, false);
                            break;
                    }

                    if (columns > 18)
                    {
                        Sobressalente_validar(rows, columns, row[column].ToString());
                    }

                    columns++;
                }

                Progresso_atualizar(total, rows);

                rows++;
            }

            Progresso_gerenciar(false);
        }

        public void Sub_grupos(DataTable dataTable, int rows)
        {
            Progresso_gerenciar(true);

            int total = dataTable.Rows.Count;

            foreach (DataRow row in dataTable.Rows)
            {
                int columns = 1;

                foreach (DataColumn column in dataTable.Columns)
                {
                    switch (columns)
                    {
                        case 1: // A - Subgrupo*
                            Campos_validar_gerenciar("Grupo ID", row[column].ToString(), rows, columns, "nivel", 8, true);
                            break;

                        case 2: // B - Subgrupo*
                            Campos_validar_gerenciar("Grupo Solution", row[column].ToString(), rows, columns, "nivel", 8, true);
                            break;

                        case 3: // C - Descrição*
                            Campos_validar_gerenciar("Descrição", row[column].ToString(), rows, columns, "char", 60, true);
                            break;

                        case 4: // D - Nível*
                            List<String> dom_nivel = new List<String> { "1", "2", "3", "4" };
                            Dominio_validar("Situação", row[column].ToString(), rows, columns, dom_nivel, true);
                            break;

                        case 5: // E - Situação*
                            List<String> dom_situacao = new List<String> { "A" };
                            Dominio_validar("Situação", row[column].ToString(), rows, columns, dom_situacao, true);
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

            Progresso_gerenciar(false);
        }
    }
}