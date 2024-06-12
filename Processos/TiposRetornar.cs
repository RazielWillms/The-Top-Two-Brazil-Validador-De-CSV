using ClosedXML.Excel;
using MathNet.Numerics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ValidarCSV
{
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

        public void Formato_date_retornar(double tipo, ref string formato)
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

            formato = formatos.ContainsKey(tipo) ? formatos[tipo] : "NULL";
        }

        public List<string> Dominio_retornar(double tipo)
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
                
            };

            return listas.ContainsKey(tipo) ? listas[tipo] : new List<string>();
        }
    }
}