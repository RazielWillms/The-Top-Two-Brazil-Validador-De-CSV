﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static ValidarCSV.TypeExtensions;

namespace ValidarCSV
{
    public static class TypeExtensions
    {

        public enum TipoCampoType
        {
            Indefinido,
            Character,
            Integer,
            Numeric,
            Date,
            //campos especiais
            DateFormat,
            Dominio,
            Nivel,
            Sintetica,
            InscricaoEstadual,
        }

        public enum Estados
        { 
            AC,
            AL,
            AP,
            AM,
            BA,
            CE,
            DF,
            ES,
            GO,
            MA,
            MT,
            MS,
            MG,
            PA,
            PB,
            PR,
            PE,
            PI,
            RJ,
            RN,
            RS,
            RO,
            RR,
            SC,
            SP,
            SE,
            TO
        }

        public static readonly Dictionary<string, (double, double)> estadosConfiguracaoIE = new Dictionary<string, (double, double)>
        {
            {"AC", (13,  0)},
            {"AL", (9,   0)},
            {"AP", ( 9,  0)},
            {"AM", ( 9,  0)},
            {"BA", ( 8,  9)},
            {"CE", ( 9,  0)},
            {"DF", (13,  0)},
            {"ES", ( 9,  0)},
            {"GO", ( 9,  0)},
            {"MA", ( 9,  0)},
            {"MT", (11,  0)},
            {"MS", ( 9,  0)},
            {"MG", (13,  0)},
            {"PA", ( 9,  0)},
            {"PB", ( 9,  0)},
            {"PR", (10,  0)},
            {"PE", ( 9,  0)},
            {"PI", ( 9,  0)},
            {"RJ", ( 9,  0)},
            {"RN", ( 9, 10)},
            {"RS", (10,  0)},
            {"RO", ( 9, 14)},
            {"RR", ( 9,  0)},
            {"SC", ( 9,  0)},
            {"SP", (12,  0)},
            {"SE", ( 9,  0)},
            {"TO", ( 9, 11)}
        };

        public static (double tam1, double tam2) Estado_config_retornar(string uf)
        {
            var estadoUFToDouble = new Dictionary<string, (double, double)>
            {
                {"AC", (13,  0)},
                {"AL", (9,   0)},
                {"AP", (9,   0)},
                {"AM", (9,   0)},
                {"BA", (8,   9)},
                {"CE", (9,   0)},
                {"DF", (13,  0)},
                {"ES", (9,   0)},
                {"GO", (9,   0)},
                {"MA", (9,   0)},
                {"MT", (11,  0)},
                {"MS", (9,   0)},
                {"MG", (13,  0)},
                {"PA", (9,   0)},
                {"PB", (9,   0)},
                {"PR", (10,  0)},
                {"PE", (9,   0)},
                {"PI", (9,   0)},
                {"RJ", (9,   0)},
                {"RN", (9,  10)},
                {"RS", (10,  0)},
                {"RO", (9,  14)},
                {"RR", (9,   0)},
                {"SC", (9,   0)},
                {"SP", (12,  0)},
                {"SE", (9,   0)},
                {"TO", (9,  11)}
            };

            return estadoUFToDouble.TryGetValue(uf, out var valores) ? valores : (0, 0);
        }


        public static readonly Dictionary<string, TipoCampoType> TipoCampoType_stringToEnum = new Dictionary<string, TipoCampoType>
        {
            { "",                   TipoCampoType.Indefinido },
            { "Character",          TipoCampoType.Character },
            { "Inteiro",            TipoCampoType.Integer },
            { "Numérico",           TipoCampoType.Numeric },
            { "Data",               TipoCampoType.Date },
            { "Data Específica",    TipoCampoType.DateFormat },
            { "Domínio",            TipoCampoType.Dominio },
            { "Nível",              TipoCampoType.Nivel },
            { "Sintética",          TipoCampoType.Sintetica },
            { "Inscrição Estadual", TipoCampoType.InscricaoEstadual }
        };

        public static void TipoCampo_string_retornar(this TipoCampoType tipoCampoType, ref string tipoCampo)
        {
            var tipoCampos = new Dictionary<TipoCampoType, string>
            {
                { TipoCampoType.Indefinido,         "" },
                { TipoCampoType.Character,          "Character" },
                { TipoCampoType.Integer,            "Inteiro" },
                { TipoCampoType.Numeric,            "Numérico" },
                { TipoCampoType.Date,               "Data" },
                { TipoCampoType.DateFormat,         "Data Específica" },
                { TipoCampoType.Dominio,            "Domínio" },
                { TipoCampoType.Nivel,              "Nível" },
                { TipoCampoType.Sintetica,          "Sintética" },
                { TipoCampoType.InscricaoEstadual,  "Inscrição Estadual" }
            };

            tipoCampo = tipoCampos.ContainsKey(tipoCampoType) ? tipoCampos[tipoCampoType] : "NULL";
        }

        public static void TipoCampo_enum_retornar(string tipoCampoType, ref TipoCampoType tipoCampo)
        {
            var tipoCampos = new Dictionary<string, TipoCampoType>
            {
                { "",                   TipoCampoType.Indefinido },
                { "Character",          TipoCampoType.Character },
                { "Inteiro",            TipoCampoType.Integer },
                { "Numérico",           TipoCampoType.Numeric },
                { "Data",               TipoCampoType.Date },
                { "Data Específica",    TipoCampoType.DateFormat },
                { "Domínio",            TipoCampoType.Dominio },
                { "Nível",              TipoCampoType.Nivel },
                { "Sintética",          TipoCampoType.Sintetica },
                { "Inscrição Estadual", TipoCampoType.InscricaoEstadual }
            };

            tipoCampo = tipoCampos.ContainsKey(tipoCampoType) ? tipoCampos[tipoCampoType] : TipoCampoType.Indefinido;
        }

        public enum CabecalhoType
        { 
            Indefinido,
            Auto,
            Sim,
            Nao,
        };

        public static readonly Dictionary<string, CabecalhoType> Cabecalho_stringToEnum = new Dictionary<string, CabecalhoType>
        {
            { "",       CabecalhoType.Indefinido},
            { "Auto",   CabecalhoType.Auto},
            { "Sim",    CabecalhoType.Sim},
            { "Não",    CabecalhoType.Nao},
        };


        public static int Indice_Cabecalho_Retornar(CabecalhoType cabecalho)
        {
            var indices = new Dictionary<CabecalhoType, int>

            {
                { CabecalhoType.Indefinido, 1 },
                { CabecalhoType.Sim,        2 },
                { CabecalhoType.Nao,        3 },
                { CabecalhoType.Auto,       4 },
            };

            return indices.ContainsKey(cabecalho) ? indices[cabecalho] : 0;
        }


        public enum LayoutType
        {
            Indefinido,
            Produtos,
            Maquinas,
            MaquinasCompleto,
            SaldosProdutosMaquinas,
            Adiantamentos,
            OrcamentoBalcao,
            OrcamentoOficina,
            Estatisticas,
            VeiculosClientes,
            ImobilizadoItens,
            ImobilizadoSaldos,
            LegadoFinanceiro,
            LegadoPagamentos,
            LegadoPedidos,
            LegadoPedidosItens,
            LegadoMovimentacao,
            Grupos,
            SubGrupos,
            Plano,
            Contas,
            Titulos,
            Enderecos,
            Prateleiras,
        }

        public static readonly Dictionary<string, LayoutType> Layout_stringToEnum = new Dictionary<string, LayoutType>
        {
            { "",                               LayoutType.Indefinido },
            { "Produtos",                       LayoutType.Produtos },
            { "Máquinas",                       LayoutType.Maquinas },
            { "Máquinas (Base + Individual)",   LayoutType.MaquinasCompleto },
            { "Saldos (Produtos ou Máquinas)",  LayoutType.SaldosProdutosMaquinas },
            { "Adiantamentos",                  LayoutType.Adiantamentos },
            { "Orçamento Balcão",               LayoutType.OrcamentoBalcao },
            { "Orçamento Oficina",              LayoutType.OrcamentoOficina },
            { "Estatísticas",                   LayoutType.Estatisticas },
            { "Veículos de Clientes",           LayoutType.VeiculosClientes },
            { "Itens Imobilizados",             LayoutType.ImobilizadoItens },
            { "Saldos Imobilizados",            LayoutType.ImobilizadoSaldos },
            { "Legado Financeiro",              LayoutType.LegadoFinanceiro },
            { "Legado Pagamentos",              LayoutType.LegadoPagamentos },
            { "Legado Pedidos",                 LayoutType.LegadoPedidos },
            { "Itens de Pedidos Legados",       LayoutType.LegadoPedidosItens },
            { "Movimentação Legada",            LayoutType.LegadoMovimentacao },
            { "Grupos",                         LayoutType.Grupos },
            { "Subgrupos",                      LayoutType.SubGrupos },
            { "Plano de Contas",                LayoutType.Plano },
            { "Contas",                         LayoutType.Contas },
            { "Títulos",                        LayoutType.Titulos },
            { "Endereços Complementares",       LayoutType.Enderecos },
            { "Prateleiras",                    LayoutType.Prateleiras },
        };

        public static void Layout_string_retornar(this LayoutType layoutType, ref string layout)
        {
            var layouts = new Dictionary<LayoutType, string>
            {
                { LayoutType.Produtos,                  "Produtos" },
                { LayoutType.Maquinas,                  "Máquinas" },
                { LayoutType.MaquinasCompleto,          "Máquinas (Base + Individual)" },
                { LayoutType.SaldosProdutosMaquinas,    "Saldos (Produtos ou Máquinas)" },
                { LayoutType.Adiantamentos,             "Adiantamentos" },
                { LayoutType.OrcamentoBalcao,           "Orçamento Balcão" },
                { LayoutType.OrcamentoOficina,          "Orçamento Oficina" },
                { LayoutType.Estatisticas,              "Estatísticas" },
                { LayoutType.VeiculosClientes,          "Veículos de Clientes" },
                { LayoutType.ImobilizadoItens,          "Itens Imobilizados" },
                { LayoutType.ImobilizadoSaldos,         "Saldos Imobilizados" },
                { LayoutType.LegadoFinanceiro,          "Legado Financeiro" },
                { LayoutType.LegadoPagamentos,          "Legado Pagamentos" },
                { LayoutType.LegadoPedidos,             "Legado Pedidos" },
                { LayoutType.LegadoPedidosItens,        "Itens de Pedidos Legados" },
                { LayoutType.LegadoMovimentacao,        "Movimentação Legada" },
                { LayoutType.Grupos,                    "Grupos" },
                { LayoutType.SubGrupos,                 "Subgrupos" },
                { LayoutType.Plano,                     "Plano de Contas" },
                { LayoutType.Contas,                    "Contas" },
                { LayoutType.Titulos,                   "Títulos" },
                { LayoutType.Enderecos,                 "Endereços Complementares" },
                { LayoutType.Prateleiras,               "Prateleiras" }
            };

            layout = layouts.ContainsKey(layoutType) ? layouts[layoutType] : "NULL";
        }

        public static void Layout_enum_retornar(string layout, ref LayoutType layoutType)
        {
            var layoutTypes = new Dictionary<string, LayoutType>
            {
                { "Produtos",                       LayoutType.Produtos },
                { "Máquinas",                       LayoutType.Maquinas },
                { "Máquinas (Base + Individual)",   LayoutType.MaquinasCompleto },
                { "Saldos (Produtos ou Máquinas)",  LayoutType.SaldosProdutosMaquinas },
                { "Adiantamentos",                  LayoutType.Adiantamentos },
                { "Orçamento Balcão",               LayoutType.OrcamentoBalcao },
                { "Orçamento Oficina",              LayoutType.OrcamentoOficina },
                { "Estatísticas",                   LayoutType.Estatisticas },
                { "Veículos de Clientes",           LayoutType.VeiculosClientes },
                { "Itens Imobilizados",             LayoutType.ImobilizadoItens },
                { "Saldos Imobilizados",            LayoutType.ImobilizadoSaldos },
                { "Legado Financeiro",              LayoutType.LegadoFinanceiro },
                { "Legado Pagamentos",              LayoutType.LegadoPagamentos },
                { "Legado Pedidos",                 LayoutType.LegadoPedidos },
                { "Itens de Pedidos Legados",       LayoutType.LegadoPedidosItens },
                { "Movimentação Legada",            LayoutType.LegadoMovimentacao },
                { "Grupos",                         LayoutType.Grupos },
                { "Subgrupos",                      LayoutType.SubGrupos },
                { "Plano de Contas",                LayoutType.Plano },
                { "Contas",                         LayoutType.Contas },
                { "Títulos",                        LayoutType.Titulos },
                { "Endereços Complementares",       LayoutType.Enderecos },
                { "Prateleiras",                    LayoutType.Prateleiras }
            };

            layoutType = layoutTypes.ContainsKey(layout) ? layoutTypes[layout] : LayoutType.Indefinido;

        }

        public enum DominioType
        {
            Nivel,
            Situacao,
            Controla_estoque,
            Importado_nacional,
            Situacao_grupos,
            Usado,
            Classe,
            Controle,
            Cliente_fornecedor,
            Tipo_operacao,
            Orcamento_situacao,
            Status,
            Novo_usado,
            Tipo_equipamento,
            Tipo_renavam_denatram,
            Especie_veiculo_renavam_denatram,
            Debito_credito,
            Chave,
            Tipo_lancamento,
            Tipo_baixa,
            Tipo_documento,
            Pagar_receber,
            Null,
            Modulo,
            Tipo,
            Pagamento,
            Forma_pagamento,
            Tipo_item,
            Tipo_movimentacao,
            Area,
            Tipo_grupo,
            Invalidos,
            Original_paralela,
            Tipo_sped,
            Curva_abc,
            Controle_estoque,
            Sintetica_analitica,
            Fisica_juridica,
            Contas_situacao,
            Tipo_contribuinte,
            Regime_tributario,
            Sim_nao,
            Estado_civil,
            Tipo_fornecedor,
            Indicador_ie,
            Cadite_situa,
            Portador,
            Xave
        }

        /*public static readonly Dictionary<string, DominioType> Dominio_stringToEnum = new Dictionary<string, DominioType>
        {
            { "Nível", DominioType.Nivel },
            { "Situação", DominioType.Situacao },
            { "Controla Estoque", DominioType.Controla_estoque },
            { "Importado/Nacional", DominioType.Importado_nacional },
            { "Situação Grupos", DominioType.Situacao_grupos },
            { "Usado", DominioType.Usado },
            { "Classe", DominioType.Classe },
            { "Controle", DominioType.Controle },
            { "Tipo Adiantamento", DominioType.Cliente_fornecedor },
            { "Tipo Operação", DominioType.Tipo_operacao },
            { "Orçamento Situação", DominioType.Orcamento_situacao },
            { "Status", DominioType.Status },
            { "Novo/Usado", DominioType.Novo_usado },
            { "Tipo Equipamento", DominioType.Tipo_equipamento },
            { "Tipo Renavam/Denatram", DominioType.Tipo_renavam_denatram },
            { "Espécie Veículo Renavam/Denatram", DominioType.Especie_veiculo_renavam_denatram },
            { "Débito/Crédito", DominioType.Debito_credito },
            { "Chave", DominioType.Chave },
            { "Tipo Lançamento", DominioType.Tipo_lancamento },
            { "Tipo Baixa", DominioType.Tipo_baixa },
            { "Tipo Documento", DominioType.Tipo_documento },
            { "Pagar/Receber", DominioType.Pagar_receber },
            { "Nulo", DominioType.Null },
            { "Módulo", DominioType.Modulo },
            { "Tipo", DominioType.Tipo },
            { "Pagamento", DominioType.Pagamento },
            { "Forma Pagamento", DominioType.Forma_pagamento },
            { "Tipo Item", DominioType.Tipo_item },
            { "Tipo Movimentação", DominioType.Tipo_movimentacao },
            { "Área", DominioType.Area },
            { "Tipo Grupo", DominioType.Tipo_grupo },
            { "Invalidos", DominioType.Invalidos },
            { "Original ou Paralela", DominioType.Original_paralela },
            { "Tipo Sped", DominioType.Tipo_sped },
            { "Curva ABC", DominioType.Curva_abc },
            { "Controle Estoque", DominioType.Controle_estoque },
            { "Sintética ou Analítica", DominioType.Sintetica_analitica },
            { "Física ou Jurídica", DominioType.Fisica_juridica },
            { "Situação", DominioType.Contas_situacao },
            { "Tipo de Contribuinte", DominioType.Tipo_contribuinte },
            { "Regime Tributário Federal", DominioType.Regime_tributario },
            { "Sim ou Não", DominioType.Sim_nao },
            { "Estado Civil", DominioType.Estado_civil },
            { "Tipo Fornecedor", DominioType.Tipo_fornecedor },
            { "Indicador IE", DominioType.Indicador_ie }
        };*/

        public static void Formato_dominio_retornar(this DominioType dominioType, ref string dominio)
        {
            var dominios = new Dictionary<DominioType, string>
            {
                { DominioType.Nivel, "Nível" },
                { DominioType.Situacao, "Situação" },
                { DominioType.Controla_estoque, "Controla Estoque" },
                { DominioType.Importado_nacional, "Importado/Nacional" },
                { DominioType.Situacao_grupos, "Situação Grupos" },
                { DominioType.Usado, "Usado" },
                { DominioType.Classe, "Classe" },
                { DominioType.Controle, "Controle" },
                { DominioType.Cliente_fornecedor, "Tipo Adiantamento" },
                { DominioType.Tipo_operacao, "Tipo Operação" },
                { DominioType.Orcamento_situacao, "Orçamento Situação" },
                { DominioType.Status, "Status" },
                { DominioType.Novo_usado, "Novo/Usado" },
                { DominioType.Tipo_equipamento, "Tipo Equipamento" },
                { DominioType.Tipo_renavam_denatram, "Tipo Renavam/Denatram" },
                { DominioType.Especie_veiculo_renavam_denatram, "Espécie Veículo Renavam/Denatram" },
                { DominioType.Debito_credito, "Débito/Crédito" },
                { DominioType.Chave, "Chave" },
                { DominioType.Tipo_lancamento, "Tipo Lançamento" },
                { DominioType.Tipo_baixa, "Tipo Baixa" },
                { DominioType.Tipo_documento, "Tipo Documento" },
                { DominioType.Pagar_receber, "Pagar/Receber" },
                { DominioType.Null, "Nulo" },
                { DominioType.Modulo, "Módulo" },
                { DominioType.Tipo, "Tipo" },
                { DominioType.Pagamento, "Pagamento" },
                { DominioType.Forma_pagamento, "Forma Pagamento" },
                { DominioType.Tipo_item, "Tipo Item" },
                { DominioType.Tipo_movimentacao, "Tipo Movimentação" },
                { DominioType.Area, "Área" },
                { DominioType.Tipo_grupo, "Tipo Grupo" },
                { DominioType.Invalidos, "Invalidos" },
                { DominioType.Original_paralela, "Original ou Paralela" },
                { DominioType.Tipo_sped, "Tipo Sped" },
                { DominioType.Curva_abc, "Curva ABC" },
                { DominioType.Controle_estoque, "Controle Estoque" },
                { DominioType.Sintetica_analitica, "Sintética ou Analítica" },
                { DominioType.Fisica_juridica, "Física ou Jurídica" },
                { DominioType.Contas_situacao, "Situação" },
                { DominioType.Tipo_contribuinte, "Tipo de Contribuinte" },
                { DominioType.Regime_tributario, "Regime Tributário Federal" },
                { DominioType.Sim_nao, "Sim ou Não" },
                { DominioType.Estado_civil, "Estado Civil" },
                { DominioType.Tipo_fornecedor, "Tipo Fornecedor" },
                { DominioType.Indicador_ie, "Indicador IE" },
                { DominioType.Cadite_situa, "Situação Cadite" },
                { DominioType.Portador, "Portador Título" },
                { DominioType.Xave, "Chave" }
            };

            dominio = dominios.ContainsKey(dominioType) ? dominios[dominioType] : "NULL";
        }

        public static void Dominio_string_retornar(string dominio, ref DominioType dominioType)
        {
            var dominioTypes = new Dictionary<string, DominioType>
            {
                { "Nível", DominioType.Nivel },
                { "Situação", DominioType.Situacao },
                { "Controla Estoque", DominioType.Controla_estoque },
                { "Importado/Nacional", DominioType.Importado_nacional },
                { "Situação Grupos", DominioType.Situacao_grupos },
                { "Usado", DominioType.Usado },
                { "Classe", DominioType.Classe },
                { "Controle", DominioType.Controle },
                { "Tipo Adiantamento", DominioType.Cliente_fornecedor },
                { "Tipo Operação", DominioType.Tipo_operacao },
                { "Orçamento Situação", DominioType.Orcamento_situacao },
                { "Status", DominioType.Status },
                { "Novo/Usado", DominioType.Novo_usado },
                { "Tipo Equipamento", DominioType.Tipo_equipamento },
                { "Tipo Renavam/Denatram", DominioType.Tipo_renavam_denatram },
                { "Espécie Veículo Renavam/Denatram", DominioType.Especie_veiculo_renavam_denatram },
                { "Débito/Crédito", DominioType.Debito_credito },
                { "Chave", DominioType.Chave },
                { "Tipo Lançamento", DominioType.Tipo_lancamento },
                { "Tipo Baixa", DominioType.Tipo_baixa },
                { "Tipo Documento", DominioType.Tipo_documento },
                { "Pagar/Receber", DominioType.Pagar_receber },
                { "Nulo", DominioType.Null },
                { "Módulo", DominioType.Modulo },
                { "Tipo", DominioType.Tipo },
                { "Pagamento", DominioType.Pagamento },
                { "Forma Pagamento", DominioType.Forma_pagamento },
                { "Tipo Item", DominioType.Tipo_item },
                { "Tipo Movimentação", DominioType.Tipo_movimentacao },
                { "Área", DominioType.Area },
                { "Tipo Grupo", DominioType.Tipo_grupo },
                { "Inválidos", DominioType.Invalidos },
                { "Original ou Paralela", DominioType.Original_paralela },
                { "Tipo Sped", DominioType.Tipo_sped },
                { "Curva ABC", DominioType.Curva_abc },
                { "Controle Estoque", DominioType.Controle_estoque },
                { "Sintética ou Analítica", DominioType.Sintetica_analitica },
                { "Física ou Jurídica", DominioType.Fisica_juridica },
                { "Situação", DominioType.Contas_situacao },
                { "Tipo de Contribuinte", DominioType.Tipo_contribuinte },
                { "Regime Tributário Federal", DominioType.Regime_tributario },
                { "Sim ou Não", DominioType.Sim_nao },
                { "Estado Civil", DominioType.Estado_civil },
                { "Tipo Fornecedor", DominioType.Tipo_fornecedor },
                { "Indicador IE", DominioType.Indicador_ie },
                { "Situação Cadite", DominioType.Cadite_situa },
                { "Portador Título", DominioType.Portador },
                { "Chave", DominioType.Xave }
            };

            dominioType = dominioTypes.ContainsKey(dominio) ? dominioTypes[dominio] : DominioType.Null;
        }

    }

    public partial class Main : Form
    {
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

        public string Formato_date_retornar(double tipo)
        {
            var formatos = new Dictionary<double, string>
            {
                { 1, "dd-MM-yyyy" },
                { 2, "yyyy-MM-dd" },
                { 3, "yyyy/MM/dd" },
                { 4, "dd/MM/yyyy" },
                { 5, "yyyy-MM-dd HH:mm:ss" },
                { 6, "dd-MM-yyyy HH:mm:ss" },
                { 7, "yyyy/MM/dd HH:mm:ss" },
                { 8, "dd/MM/yyyy HH:mm:ss" }
            };
            return _ = formatos.ContainsKey(tipo) ? formatos[tipo] : "NULL";
        }

        public List<string> Dominio_lista_retornar(double tipo)
        {
            var listas = new Dictionary<double, List<string>>
            {
                { 1, new List<string> { "1", "2", "3", "4" } }, //dom_nivel
                { 2, new List<string> { "A" } }, //dom_situacao
                { 3, new List<String> { "S", "N" } }, //controla_estoque 
                { 4, new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8" } }, //dom_importado_nacional
                { 5, new List<string> { "A", "I" } }, //dom_situacao
                { 6, new List<string> { "1", "0" } }, //dom_usado
                { 7, new List<string> { "N", "B" } }, //dom_classe
                { 8, new List<string> { "I" } }, //dom_controle
                { 9, new List<String> { "C", "F" } }, //dom_tipo_adiantamento
                { 10, new List<String> { "V", "S", "E", "C", "D" } }, //dom_tipo_operacao
                { 11, new List<String> { "A", "F" } }, //dom_orcamento_situacao
                { 12, new List<String> { "A", "P", "C", "F", "B", "S", "X", "Y" } }, //dom_status
                { 13, new List<String> { "N", "U" } }, //dom_novo_usado
                { 14, new List<String> { "#", "J", "8", "4", "A", "5", "N", "C", "R", "D", "2", "L", "K", "P", "H", "V", "I", "3", "S", "6", "M", "O", "9", "Z", "B", "U", "F", "7", "Y", "T", "G", "Q", "1", "E", "X" } }, //dom_tipo_equipamento
                { 15, new List<String> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26" } }, //dom_tipo_renavam_denatram
                { 16, new List<String> { "0", "1", "2", "3", "4", "5", "6" } }, //dom_especie_veiculo_renavam_denatram
                { 17, new List<String> { "D", "C" } }, //dom_debito_credito
                { 18, new List<String> { "G", "C" } }, //dom_chave
                { 19, new List<String> { "A", "T", "I" } }, //dom_tipo_lancamento
                { 20, new List<String> { "B", "T" } }, //dom_tipo_baixa
                { 21, new List<String> { "#", "C", "T", "A" } }, //dom_tipo_documento
                { 22, new List<String> { "P", "R" } }, //dom_pagar_receber
                { 23, new List<String> { "", "null", "NULL" } }, //dom_null
                { 24, new List<String> { "5", "17" } }, //dom_modulo
                { 25, new List<String> { "O", "P" } }, //dom_tipo
                { 26, new List<String> { "V", "P" } }, //dom_pagamento
                { 27, new List<String> { "A", "2", "4", "5", "0", "1", "6", "3", "F", "9", "8" } }, //dom_forma_pagamento
                { 28, new List<String> { "SP", "P", "ST" } }, //dom_tipo_item
                { 29, new List<String> { "S", "E" } }, //dom_tipo_movimentacao
                { 30, new List<String> { "1" } }, //dom_area
                { 31, new List<string> { "E" } }, //dom_tipo
                { 32, new List<string> { "#", "0", "", " ", "null", "NULL" } }, //dom_invalidos
                { 33, new List<string> { "O", "P"} }, //dom_original_paralela
                { 34, new List<string> { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "99" } }, //dom_tipo_sped
                { 35, new List<string> { "A", "B", "C"} }, //dom_curva_abc
                { 36, new List<string> { "N", "I"} }, //dom_controle_estoque
                { 37, new List<string> { "S", "A"} }, //dom_sintetica_analitica
                { 38, new List<string> { "F", "J"} }, //dom_fisica_juridica
                { 39, new List<string> { "A", "I", "B", "C", "L", "M", "R", "S"} }, //dom_contas_situacao
                { 40, new List<string> { "N", "F", "E", "R", "I", "S", "P", "A"} }, //dom_tipo_contribuinte
                { 41, new List<string> { "0", "1", "2"} }, //dom_regime_tributario
                { 42, new List<string> { "S", "N"} }, //dom_sim_nao
                { 43, new List<string> { "S", "C", "V", "J"} }, //dom_estado_civil
                { 44, new List<string> { "C", "I", "T", "F", ""} }, //dom_tipo_fornecedor
                { 45, new List<string> { "0", "1", "2", "9"} }, //dom_indicador_ie
                { 46, new List<string> { "A", "I", "C" } }, //dom_cadite_situa
                { 47, new List<string> { "0", "1", "2", "4", "P" } }, //dom_portador
                { 48, new List<string> { "0", "1" } }, //dom_xave
                
            };

            return listas.ContainsKey(tipo) ? listas[tipo] : new List<string>();
        }

        public double Dominio_retornar(DominioType dominioType)
        {
            var dominioTypeToDouble = new Dictionary<DominioType, double>
            {
                { DominioType.Nivel, 1 },
                { DominioType.Situacao_grupos, 2 },
                { DominioType.Controla_estoque, 3 },
                { DominioType.Importado_nacional, 4 },
                { DominioType.Situacao, 5 },
                { DominioType.Usado, 6 },
                { DominioType.Classe, 7 },
                { DominioType.Controle, 8 },
                { DominioType.Cliente_fornecedor, 9 },
                { DominioType.Tipo_operacao, 10 },
                { DominioType.Orcamento_situacao, 11 },
                { DominioType.Status, 12 },
                { DominioType.Novo_usado, 13 },
                { DominioType.Tipo_equipamento, 14 },
                { DominioType.Tipo_renavam_denatram, 15 },
                { DominioType.Especie_veiculo_renavam_denatram, 16 },
                { DominioType.Debito_credito, 17 },
                { DominioType.Chave, 18 },
                { DominioType.Tipo_lancamento, 19 },
                { DominioType.Tipo_baixa, 20 },
                { DominioType.Tipo_documento, 21 },
                { DominioType.Pagar_receber, 22 },
                { DominioType.Null, 23 },
                { DominioType.Modulo, 24 },
                { DominioType.Tipo, 25 },
                { DominioType.Pagamento, 26 },
                { DominioType.Forma_pagamento, 27 },
                { DominioType.Tipo_item, 28 },
                { DominioType.Tipo_movimentacao, 29 },
                { DominioType.Area, 30 },
                { DominioType.Tipo_grupo, 31 },
                { DominioType.Invalidos, 32 },
                { DominioType.Original_paralela, 33 },
                { DominioType.Tipo_sped, 34 },
                { DominioType.Curva_abc, 35 },
                { DominioType.Controle_estoque, 36 },
                { DominioType.Sintetica_analitica, 37 },
                { DominioType.Fisica_juridica, 38 },
                { DominioType.Contas_situacao, 39 },
                { DominioType.Tipo_contribuinte, 40 },
                { DominioType.Regime_tributario, 41 },
                { DominioType.Sim_nao, 42 },
                { DominioType.Estado_civil, 43 },
                { DominioType.Tipo_fornecedor, 44 },
                { DominioType.Indicador_ie, 45 },
                { DominioType.Cadite_situa, 46 },
                { DominioType.Portador, 47 },
                { DominioType.Xave, 48 }
            };

            double dominio;
            if (dominioTypeToDouble.TryGetValue(dominioType, out double value))
            {
                dominio = value;
            }
            else
            {
                dominio = dominioTypeToDouble[DominioType.Null];
            }

            return dominio;
        }
    }
}