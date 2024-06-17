using ClosedXML.Excel;
using MathNet.Numerics;
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System.Globalization;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        public bool Obrigatorio_validar(string campo, string tipo, ref string mensagem)
        {
            if (tipo == "integer" || tipo == "numeric")
            {
                if (!Int32.TryParse(campo, out _) && !decimal.TryParse(campo, out _))
                {
                    mensagem = "Formato inválido";
                }
                else if ((Int32.TryParse(campo, out int valorInteiro) && valorInteiro <= 0) || (decimal.TryParse(campo, out decimal valorDecimal) && valorDecimal <= 0))
                {
                    mensagem = "Deve ser maior que zero";
                }
            }

            string[] invalidos = { "#", "0", "", "null", "NULL" };
            if (string.IsNullOrEmpty(mensagem) && (invalidos.Contains(campo.Trim())))
            {
                mensagem = "Campo obrigatório";
            }

            if (string.IsNullOrEmpty(mensagem) && string.IsNullOrEmpty(campo))
            {
                mensagem = "Campo está vazio";
            }

            if (!string.IsNullOrEmpty(mensagem))
            {
                return true;
            }

            return false;
        }

        public void Campos_validar_gerenciar(string tabela, string campo, int linha, int coluna, string tipo, double tamanho_formato, Boolean obrigatorio)
        {
            string mensagem = string.Empty;

            if (obrigatorio && Obrigatorio_validar(campo, tipo, ref mensagem))
            {
                Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                return;
            }

            List<string> domVazio = new List<string> { "", "#", "0", "null", "NULL" };
            if (!obrigatorio && domVazio.Contains(campo))
            {
                return;
            }

            bool valido = true;

            switch (tipo.ToLower())
            {
                //campos padrão
                case "char":
                    Char_validar(campo, tamanho_formato, ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                case "numeric":
                    Numeric_validar(campo.Trim(), tamanho_formato, ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                case "date":
                    Date_validar(campo.Trim(), ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                case "date_format":
                    Date_formato_validar(campo.Trim(), Formato_date_retornar(tamanho_formato), ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                case "integer":
                    Integer_validar(campo, tamanho_formato, ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                //Campos 'especiais'
                case "nivel": //Grupos e Subgrupos
                    Nivel_validar(campo.Trim(), tamanho_formato, ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                case "dominio": //provindos de enum do genexus
                    Dominio_validar(campo, tamanho_formato, obrigatorio, ref mensagem, ref valido);
                    if (!valido)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                    }
                    break;

                default:
                    Registro_adicionar(tabela, linha, coluna, campo, "Validação falhou, conferir manualmente");
                    break;
            }
        }

        private void Char_validar(string campo, double tamanho_formato, ref string mensagem, ref bool valido)
        {
            if (campo.Length > tamanho_formato)
            {
                valido = false;
                mensagem = "Excede " + tamanho_formato.ToString() + " caracteres";
            }
        }

        private void Integer_validar(string campo, double tamanho_formato, ref string mensagem, ref bool valido)
        {
            campo = campo.Replace(".", "");
            if (campo == "0" || campo.Trim() == "")
            {
                return;
            }

            if (campo.Length > tamanho_formato || !int.TryParse(campo, out _))
            {
                valido = false;
                mensagem = "Deve ser um número inteiro e conter até " + tamanho_formato + " dígitos";
            }
        }

        private void Numeric_validar(string valor, double tamanho_formato, ref string mensagem_erro, ref bool valido)
        {
            valor = valor.Replace(".", "");
            if (valor != "0" && valor.Trim() != "")
            {
                return;
            }

            int precisao = (int)Math.Truncate(tamanho_formato);
            double parteDecimal = (tamanho_formato - precisao).Round(1);
            //parteDecimal = parteDecimal.Round(1);
            int escala = (int)(parteDecimal * 10);            

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

        static void Date_validar(string data, ref string mensagem, ref bool valido) //Válido qualquer formato, já que pode ser escolhido no -converte
        {
            if (string.IsNullOrWhiteSpace(data) || data.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            string[] formatos = { "dd-MM-yyyy", "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy" };
            valido = DateTime.TryParseExact(data, formatos, null, System.Globalization.DateTimeStyles.None, out _);

            if (!valido)
            {
                string[] formatos_invalidos = { "d-M-yyyy", "yyyy-M-d", "yyyy/M/d", "d/M/yyyy", "d-M-yy", "yy-M-d", "yy/M/d", "d/M/yy" };
                bool isValid = DateTime.TryParseExact(data, formatos_invalidos, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

                if (isValid)
                {
                    mensagem = "Inválido: informe os meses e dias com zeros à esquerda e utilize 4 dígitos para o ano";
                }
                else 
                {
                    mensagem = "Deve estar em um formato de data válido";
                }
            } 
        }

        private void Date_formato_validar(string data, string formato, ref string mensagem, ref bool valido) //Valida formato específico, quando necessário ficar como indicado no layout
        {
            if (string.IsNullOrWhiteSpace(data) || data.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            valido = DateTime.TryParseExact(data, formato, null, System.Globalization.DateTimeStyles.None, out _);

            if (!valido)
            {
                mensagem = "Deve estar em um formato de data válido, conforme layout: " + formato;
            }
        }

        private void Nivel_mensagem_retornar(string campo, double tamanho_formato, string mensagem, ref string mensagem_completa, ref bool valido)
        {
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
        }

        private void Dominio_validar(string campo, double tamanho_formato, bool obrigatorio, ref string mensagem, ref bool valido)
        {
            mensagem = string.Empty;
            List<string> dominio = Dominio_lista_retornar(tamanho_formato);
            List<string> dominioExtendido = new List<string>(dominio) { "", "null", "NULL" };

            if (!dominioExtendido.Contains(campo.Trim()))
            {
                valido = false;
                string opcoes = String.Join(", ", dominio);

                if (obrigatorio)
                {
                    mensagem = $"Deve estar entre as opções: {opcoes}";
                }
                else
                {
                    mensagem = $"Deve estar entre as opções: {opcoes} ou vazio.";
                }
            }
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