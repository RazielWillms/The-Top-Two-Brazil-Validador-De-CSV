using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using MathNet.Numerics;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        public bool Obrigatorio_validar(string tabela, string campo, int linha, int coluna, string tipo)
        {

            if (campo.Trim() == "#" || campo.Trim() == "0" || campo.Trim() == "" || campo.Trim() == "null" || campo.Trim() == "NULL")
            {
                Registro_adicionar(tabela, linha, coluna, campo, "Campo obrigatório");
                return true;
            }

            if (string.IsNullOrEmpty(campo))
            {
                Registro_adicionar(tabela, linha, coluna, campo, "Campo está vazio");
                return true;
            }

            if (tipo == "integer")
            {
                if (Int32.TryParse(campo, out int valorInteiro))
                {
                    if (valorInteiro <= 0)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve ser maior que zero");
                        return true;
                    }
                }
                else
                {
                    Registro_adicionar(tabela, linha, coluna, campo, "Formato inválido");
                    return true;
                }
            }

            if (tipo == "numeric")
            {
                if (decimal.TryParse(campo, out decimal valorDecimal))
                {
                    if (valorDecimal <= 0)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve ser maior que zero");
                        return true;
                    }
                }
                else
                {
                    Registro_adicionar(tabela, linha, coluna, campo, "Formato inválido");
                    return true;
                }
            }

            return false;
        }

        public void Dominio_validar(string tabela, string campo, int linha, int coluna, List<String> dominio, Boolean obrigatorio)
        {
            if (obrigatorio)
            {
                if (Obrigatorio_validar(tabela, campo, linha, coluna, "N"))
                {
                    return;
                }
            }

            string opcoes = String.Join(", ", dominio);

            dominio.Add("");
            dominio.Add("null");
            dominio.Add("NULL");
            
            if (!dominio.Contains(campo.Trim()))
            {
                if (obrigatorio)
                {
                    Registro_adicionar(tabela, linha, coluna, campo, "Deve estar entre as opções: " + opcoes);
                }
                else 
                {
                    Registro_adicionar(tabela, linha, coluna, campo, "Deve estar entre as opções: " + opcoes + " ou vazio.");
                }
            }
        }

        public void Campos_validar_gerenciar(string tabela, string campo, int linha, int coluna, string tipo, double tamanho, Boolean obrigatorio)
        {
            if (obrigatorio)
            {
                if (Obrigatorio_validar(tabela, campo, linha, coluna, tipo))
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
                        Registro_adicionar(tabela, linha, coluna, campo, "Excede " + tamanho.ToString() + " caracter");
                    }
                    break;

                case "numeric":
                    if (campo != "0" && campo.Trim() != "")
                    {
                        int parteInteira = (int)Math.Truncate(tamanho);
                        double parteDecimal = tamanho - parteInteira;
                        parteDecimal = parteDecimal.Round(1);
                        int parteFracionaria = (int)(parteDecimal * 10);
                        string mensagem = string.Empty;
                        bool valido = true;

                        Numeric_validar(campo.Trim(), parteInteira, parteFracionaria, ref mensagem, ref valido);
                        if (!valido)
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                            //Registro_adicionar(tabela, linha, coluna, campo, "Deve estar no formato numérico: '" + tamanho.ToString().Replace('.', ',') + "'");
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
                    Formato_retornar(tamanho, ref formato);

                    if (!Date_formato_validar(campo.Trim(), formato))
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve estar em um formato de data válido, conforme layout: " + formato);
                    }
                    break;

                case "integer":
                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho || !int.TryParse(campo, out _))
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, "Deve ser um número inteiro e conter até " + tamanho + " dígitos");
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

        /*private bool Numeric_validar(string valor, int precisao, int escala, string mensagem_erro)
        {
            if (valor == null || valor == "" || valor == "null" || valor == "NULL")
            {
                return true;
            }

            string pattern = @"^\d{1," + precisao.ToString().Trim() + @"}(,\d{1," + escala.ToString().Trim() + "})?$";
            return Regex.IsMatch(valor, pattern);
        }
         */

        static bool Date_validar(string data) //Válido qualquer formato, já que pode ser escolhido no -converte
        {
            if (data == null || data == "" || data == "null" || data == "NULL")
            {
                return true;
            }

            string[] formatos = { "dd-MM-yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy" };
            return DateTime.TryParseExact(data, formatos, null, System.Globalization.DateTimeStyles.None, out _);
        }

        private void Formato_retornar(double tipo, ref string formato) //Verificar qual formato para passar por parâmetro
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

        private bool Date_formato_validar(string data, string formato) //Valida formato específico, quando necessário ficar como indicado no layout
        {
            if (data == null || data == "" || data == "null" || data == "NULL")
            {
                return true;
            }

            return DateTime.TryParseExact(data, formato, null, System.Globalization.DateTimeStyles.None, out _);
        }

        private void Sobressalente_validar(int rows, int columns, string campo)
        {
            if (string.IsNullOrEmpty(campo))
            {
                return;
            }
            else 
            {
                Registro_adicionar("Erro genérico", rows, columns, campo, "Excedeu o número de colunas");
            }
        }
    }
}