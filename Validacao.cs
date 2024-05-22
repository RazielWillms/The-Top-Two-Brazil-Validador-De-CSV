using ClosedXML.Excel;
using MathNet.Numerics;
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

            string opcoes;
            if (!dominio.Contains(campo.Trim()))
            {
                opcoes = String.Join(", ", dominio);
                Registro_adicionar(tabela, linha, coluna, campo, "Deve estar entre as opções: " + opcoes);
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

                        if (!Numeric_validar(campo.Trim(), parteInteira, parteFracionaria))
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, "Deve estar no formato numérico: '" + tamanho.ToString().Replace('.', ',') + "'");
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

        private bool Numeric_validar(string valor, int precisao, int escala)
        {
            if (valor == null || valor == "" || valor == "null" || valor == "NULL")
            {
                return true;
            }

            string pattern = @"^\d{1," + precisao.ToString().Trim() + @"}(,\d{1," + escala.ToString().Trim() + "})?$";
            return Regex.IsMatch(valor, pattern);
        }

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
    }
}