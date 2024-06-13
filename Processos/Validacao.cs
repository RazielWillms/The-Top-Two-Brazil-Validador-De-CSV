﻿using ClosedXML.Excel;
using MathNet.Numerics;
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        private static int LevenshteinDistance(string string1, string string2)
        {
            if (string.IsNullOrEmpty(string1))
            {
                return string.IsNullOrEmpty(string2) ? 0 : string2.Length;
            }

            if (string.IsNullOrEmpty(string2))
            {
                return string1.Length;
            }

            int[,] d = new int[string1.Length + 1, string2.Length + 1];

            for (int i = 0; i <= string1.Length; i++)
            {
                d[i, 0] = i;
            }

            for (int j = 0; j <= string2.Length; j++)
            {
                d[0, j] = j;
            }

            for (int i = 1; i <= string1.Length; i++)
            {
                for (int j = 1; j <= string2.Length; j++)
                {
                    int cost = (string2[j - 1] == string1[i - 1]) ? 0 : 1;

                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }

            return d[string1.Length, string2.Length];
        }

        public static bool Similar_validar(string string1, string string2, int limite_diferenca)
        {
            int distance = LevenshteinDistance(string1, string2);
            return distance <= limite_diferenca;
        }

        public bool Obrigatorio_validar(string tabela, string campo, int linha, int coluna, string tipo)
        {
            string mensagemErro = string.Empty;

            if (tipo == "integer" || tipo == "numeric")
            {
                if (!Int32.TryParse(campo, out _) && !decimal.TryParse(campo, out _))
                {
                    mensagemErro = "Formato inválido";
                }
                else if ((Int32.TryParse(campo, out int valorInteiro) && valorInteiro <= 0) || (decimal.TryParse(campo, out decimal valorDecimal) && valorDecimal <= 0))
                {
                    mensagemErro = "Deve ser maior que zero";
                }
            }

            string[] invalidos = { "#", "0", "", "null", "NULL" };
            if (string.IsNullOrEmpty(mensagemErro) && (invalidos.Contains(campo.Trim())))
            {
                mensagemErro = "Campo obrigatório";
            }

            if (string.IsNullOrEmpty(mensagemErro) && string.IsNullOrEmpty(campo))
            {
                mensagemErro = "Campo está vazio";
            }

            if (!string.IsNullOrEmpty(mensagemErro))
            {
                Registro_adicionar(tabela, linha, coluna, campo, mensagemErro);
                return true;
            }

            return false;
        }

        public void Campos_validar_gerenciar(string tabela, string campo, int linha, int coluna, string tipo, double tamanho_formato, Boolean obrigatorio)
        {            

            if (obrigatorio && Obrigatorio_validar(tabela, campo, linha, coluna, tipo))
            {
                return;
            }

            List<string> domVazio = new List<string> { "", "#", "0", "null", "NULL" };
            if (!obrigatorio && domVazio.Contains(campo))
            {
                return;
            }

            string mensagem = string.Empty;
            bool valido = true;

            switch (tipo.ToLower())
            {
                //campos padrão
                case "char":
                    if (campo.Length > tamanho_formato)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Excede " + tamanho_formato.ToString() + " caracteres");
                    }
                    break;

                case "numeric":
                    campo = campo.Replace(".", "");
                    if (campo != "0" && campo.Trim() != "")
                    {
                        int parteInteira = (int)Math.Truncate(tamanho_formato);
                        double parteDecimal = (tamanho_formato - parteInteira).Round(1);
                        //parteDecimal = parteDecimal.Round(1);
                        int parteFracionaria = (int)(parteDecimal * 10);

                        Numeric_validar(campo.Trim(), parteInteira, parteFracionaria, ref mensagem, ref valido);
                        if (!valido)
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                        }
                    }
                    break;

                case "date":
                    if (!Date_validar(campo.Trim()))
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve estar em um formato de data válido");
                    }
                    break;

                case "date_format":
                    string formato = string.Empty;
                    Formato_date_retornar(tamanho_formato, ref formato);

                    if (!Date_formato_validar(campo.Trim(), formato))
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve estar em um formato de data válido, conforme layout: " + formato);
                    }
                    break;

                case "integer":
                    campo = campo.Replace(".", "");
                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho_formato || !int.TryParse(campo, out _))
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, "Deve ser um número inteiro e conter até " + tamanho_formato + " dígitos");
                        }
                    }
                    break;

                case "nivel":
                    Nivel_validar(campo.Trim(), ref mensagem, ref valido);

                    string mensagem_completa = string.Empty;
                    int tamanho_nivel = (int.Parse(NiveisCombo.Text.Substring(0, 1)) * 2);

                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho_formato || !int.TryParse(campo, out _))
                        {
                            mensagem_completa = "Deve ser um número inteiro e conter até " + tamanho_nivel.ToString() + " dígitos. ";
                            valido = false;
                        }
                    }
                    mensagem_completa += mensagem;

                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem_completa);
                    }                    
                    break;

                case "dominio":
                    List<string> dominio = Dominio_lista_retornar(tamanho_formato);
                    List<string> dominioExtendido = new List<string>(dominio) { "", "null", "NULL" };

                    if (!dominioExtendido.Contains(campo.Trim()))
                    {
                        string opcoes = String.Join(", ", dominio);

                        if (obrigatorio)
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, $"Deve estar entre as opções: {opcoes}");
                        }
                        else
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, $"Deve estar entre as opções: {opcoes} ou vazio.");
                        }
                    }
                    break;

                default:
                    Registro_adicionar(tabela, linha, coluna, campo, "Validação falhou, conferir manualmente");
                    break;
            }
        }

        private void Numeric_validar(string valor, int precisao, int escala, ref string mensagem_erro, ref bool valido)
        {
            mensagem_erro = string.Empty;

            if (string.IsNullOrEmpty(valor) || valor.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                return; 
            }

            string[] partes = valor.Split(',');

            if (partes[0].Length > precisao && partes.Length > 1 && partes[1].Length > escala)
            {
                mensagem_erro = $"Erro de precisão e escala: a parte inteira tem mais de {precisao} dígitos e a parte decimal tem mais de {escala} dígitos. ";
                valido = false;
                return;
            }

            if (partes[0].Length > precisao)
            {
                mensagem_erro = $"Erro de precisão: a parte inteira tem mais de {precisao} dígitos.";
                valido = false;
                return;
            }

            if (partes.Length > 1 && partes[1].Length > escala)
            {
                mensagem_erro = $"Erro de escala: a parte decimal tem mais de {escala} dígitos.";
                valido = false;
                return;
            }

            string pattern = @"^\d{1," + precisao.ToString().Trim() + @"}(,\d{1," + escala.ToString().Trim() + "})?$";
            if (!Regex.IsMatch(valor, pattern))
            {
                mensagem_erro = "Erro de formato: o valor não corresponde ao formato esperado. " + precisao.ToString() + "," + escala.ToString();
                valido = false;
                return;
            }
        }

        static bool Date_validar(string data) //Válido qualquer formato, já que pode ser escolhido no -converte
        {
            if (string.IsNullOrWhiteSpace(data) || data.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string[] formatos = { "dd-MM-yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy" };
            return DateTime.TryParseExact(data, formatos, null, System.Globalization.DateTimeStyles.None, out _);
        }

        private bool Date_formato_validar(string data, string formato) //Valida formato específico, quando necessário ficar como indicado no layout
        {
            if (string.IsNullOrWhiteSpace(data) || data.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return DateTime.TryParseExact(data, formato, null, System.Globalization.DateTimeStyles.None, out _);
        }

        private void Sobressalente_validar(int rows, int columns, string campo)
        {
            string[] invalidos = { "#", "0", "", "null", "NULL" };
            if (!string.IsNullOrEmpty(campo) || !invalidos.Contains(campo.Trim()))
            {
                Registro_adicionar("Erro genérico", rows, columns, campo, "Excedeu o número de colunas");
            }
        }
    }
}