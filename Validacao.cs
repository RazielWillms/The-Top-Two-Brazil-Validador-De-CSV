using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using MathNet.Numerics;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {
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

        public void Dominio_validar(string tabela, string campo, int linha, int coluna, List<String> dominio, Boolean obrigatorio)
        {
            if (obrigatorio && Obrigatorio_validar(tabela, campo, linha, coluna, "NULL"))
            {
                return;
            }

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
        }

        public void Campos_validar_gerenciar(string tabela, string campo, int linha, int coluna, string tipo, double tamanho, Boolean obrigatorio)
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
                    if (campo.Length > tamanho)
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Excede " + tamanho.ToString() + " caracter");
                    }
                    break;

                case "numeric":
                    campo = campo.Replace(".", "");
                    if (campo != "0" && campo.Trim() != "")
                    {
                        int parteInteira = (int)Math.Truncate(tamanho);
                        double parteDecimal = (tamanho - parteInteira).Round(1);
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
                    Formato_retornar(tamanho, ref formato);

                    if (!Date_formato_validar(campo.Trim(), formato))
                    {
                        Registro_adicionar(tabela, linha, coluna, campo, "Deve estar em um formato de data válido, conforme layout: " + formato);
                    }
                    break;

                case "integer":
                    campo = campo.Replace(".", "");
                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho || !int.TryParse(campo, out _))
                        {
                            Registro_adicionar(tabela, linha, coluna, campo, "Deve ser um número inteiro e conter até " + tamanho + " dígitos");
                        }
                    }
                    break;

                case "nivel":
                    Nivel_validar(campo.Trim(), ref mensagem, ref valido);

                    string mensagem_completa = string.Empty;
                    int tamanho_nivel = (int.Parse(NiveisCombo.Text.Substring(0, 1)) * 2);

                    if (campo != "0" && campo.Trim() != "")
                    {
                        if (campo.Length > tamanho || !int.TryParse(campo, out _))
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

        private void Formato_retornar(double tipo, ref string formato)
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
            if (!string.IsNullOrEmpty(campo))
            {
                Registro_adicionar("Erro genérico", rows, columns, campo, "Excedeu o número de colunas");
            }
        }

        /*private void Nivel_validar(string campo, ref string mensagem, ref bool valido)
        {
            mensagem = string.Empty;
            valido = false;

            if (campo.Contains('.'))
            {
                mensagem = "Não deve conter pontuação";
                valido = false;
                return;
            }
            
            int tamanho_nivel = (int.Parse(NiveisCombo.Text.Substring(0,1)) * 2);
            
            if (tamanho_nivel != campo.Length) 
            {
                mensagem = "campo possui " + campo.Length.ToString() + " dígitos, o nível espera " + tamanho_nivel.ToString();
                valido = false;
                return;
            }

            if (layouts.Text == "Grupos")
            {
                if (tamanho_nivel == 8)
                {
                    if (campo.Substring(2, 6) == "000000" && campo.Substring(0, 2) != "00")
                    {
                        valido = true;
                    }
                    else
                    {
                        mensagem = "Deve ser informado um Grupo (ex:99000000)";
                    }
                }
                else if (tamanho_nivel == 6)
                {
                    if (campo.Substring(2, 4) == "0000" && campo.Substring(0, 2) != "00")
                    {
                        valido = true;
                    }
                    else
                    {
                        mensagem = "Deve ser informado um Grupo (ex:990000)";
                    }
                }
                else if (tamanho_nivel == 4)
                {
                    if (campo.Substring(2, 2) == "00" && campo.Substring(0, 2) != "00")
                    {
                        valido = true;
                    }
                    else
                    {
                        mensagem = "Deve ser informado um Grupo (ex:9900)";
                    }
                }
            }
            else
            {
                switch (NivelCombo.Text)
                {
                    case "SubGrupo":
                        if (tamanho_nivel == 8)
                        {
                            if (campo.Substring(4, 4) == "0000" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00")
                            {
                                valido = true;
                            }
                            else
                            {
                                mensagem = "Deve ser informado um SubGrupo (ex:99990000)";
                            }
                        }
                        else if (tamanho_nivel == 6)
                        {
                            if (campo.Substring(4, 2) == "00" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00")
                            {
                                valido = true;
                            }
                            else
                            {
                                mensagem = "Deve ser informado um SubGrupo (ex:999900)";
                            }
                        }
                        else if (tamanho_nivel == 4)
                        {
                            if (campo.Substring(4, 2) == "00" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00")
                            {
                                mensagem = "Deve ser informado um SubGrupo (ex:9999)";
                            }
                            else
                            {
                                valido = true;
                            }
                        }
                        break;

                    case "Segmento":
                        if (tamanho_nivel == 8)
                        {
                            if (campo.Substring(6, 2) == "00" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00" && campo.Substring(4, 2) != "00")
                            {
                                valido = true;
                            }
                            else
                            {
                                mensagem = "Deve ser informado um Segmento (ex:99999900)";
                            }
                        }
                        else if (tamanho_nivel == 6)
                        {
                            if (campo.Substring(4, 2) == "00" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00" && campo.Substring(4, 2) != "00")
                            {
                                mensagem = "Deve ser informado um Segmento (ex:999999)";
                            }
                            else
                            {
                                valido = true;
                            }
                        }
                        else
                        {
                            mensagem = "Segmento não é válido para Subgrupo de " + NivelCombo.Text + "níveis.";
                        }
                        break;

                    case "SubSegmento":

                        if (tamanho_nivel == 8)
                        {
                            if (campo.Substring(6, 2) == "00" && campo.Substring(0, 2) != "00" && campo.Substring(2, 2) != "00" && campo.Substring(4, 2) != "00")
                            {
                                mensagem = "Deve ser informado um SubSegmento (ex:99999999)";
                            }
                            else
                            {
                                valido = true;
                            }
                        }
                        else
                        {
                            mensagem = "SubSegmento não é válido para Subgrupo de " + NivelCombo.Text + "níveis.";
                        }
                        break;

                    default:
                        mensagem = "Nível desconhecido";
                        break;
                }
            }
        }*/

        
        private void Nivel_validar(string campo, ref string mensagem, ref bool valido)
        {
            mensagem = string.Empty;
            valido = false;

            if (campo.Contains('.') || campo.Contains(','))
            {
                mensagem = "Não deve conter pontuação";
                return;
            }

            if (!Int32.TryParse(campo, out _) && !decimal.TryParse(campo, out _))
            {
                mensagem = "Formato inválido";
                return;
            }

            int tamanho_nivel = int.Parse(NiveisCombo.Text.Substring(0, 1)) * 2;

            if (tamanho_nivel != campo.Length)
            {
                mensagem = $"Campo possui {campo.Length} dígitos, o nível espera {tamanho_nivel}";
                return;
            }

            if (layouts.Text == "Grupos")
            {
                valido = ValidarGrupo(campo, tamanho_nivel, ref mensagem);
            }
            else
            {
                valido = ValidarSubNivel(campo, tamanho_nivel, ref mensagem, NivelCombo.Text);
            }
        }

        private bool ValidarGrupo(string campo, int tamanho_nivel, ref string mensagem)
        {
            switch (tamanho_nivel)
            {
                case 8:
                    if (campo.Substring(2, 6) != "000000" || campo.Substring(0, 2) == "00")
                    {
                        mensagem = "Deve ser informado um Grupo (ex: 09000000)";
                        return false;
                    }
                    return true;

                case 6:
                    if (campo.Substring(2, 4) != "0000" || campo.Substring(0, 2) == "00")
                    {
                        mensagem = "Deve ser informado um Grupo (ex: 090000)";
                        return false;
                    }
                    return true;

                case 4:
                    if (campo.Substring(2, 2) != "00" || campo.Substring(0, 2) == "00")
                    {
                        mensagem = "Deve ser informado um Grupo (ex: 0900)";
                        return false;
                    }
                    return true;

                default:
                    mensagem = "Tamanho de nível inválido para Grupos";
                    return false;
            }
        }

        private bool ValidarSubNivel(string campo, int tamanho_nivel, ref string mensagem, string nivel)
        {
            switch (nivel)
            {
                case "SubGrupo":
                    return ValidarSubGrupo(campo, tamanho_nivel, ref mensagem);

                case "Segmento":
                    return ValidarSegmento(campo, tamanho_nivel, ref mensagem);

                case "SubSegmento":
                    return ValidarSubSegmento(campo, tamanho_nivel, ref mensagem);

                default:
                    mensagem = "Nível desconhecido";
                    return false;
            }
        }

        private bool ValidarSubGrupo(string campo, int tamanho_nivel, ref string mensagem)
        {
            switch (tamanho_nivel)
            {
                case 8:
                    if (campo.Substring(4, 4) != "0000" || campo.Substring(0, 2) == "00" || campo.Substring(2, 2) == "00")
                    {
                        mensagem = "Deve ser informado um SubGrupo (ex: 09090000)";
                        return false;
                    }
                    return true;

                case 6:
                    if (campo.Substring(4, 2) != "00" || campo.Substring(0, 2) == "00" || campo.Substring(2, 2) == "00")
                    {
                        mensagem = "Deve ser informado um SubGrupo (ex: 090900)";
                        return false;
                    }
                    return true;

                case 4:
                    if (campo.Substring(2, 2) == "00" || campo.Substring(0, 2) == "00")
                    {
                        mensagem = "Deve ser informado um SubGrupo (ex: 0909)";
                        return false;
                    }
                    return true;

                default:
                    mensagem = "Tamanho de nível inválido para SubGrupo";
                    return false;
            }
        }

        private bool ValidarSegmento(string campo, int tamanho_nivel, ref string mensagem)
        {
            if (tamanho_nivel == 8)
            {
                if (campo.Substring(6, 2) != "00" || campo.Substring(0, 2) == "00" || campo.Substring(2, 2) == "00" || campo.Substring(4, 2) == "00")
                {
                    mensagem = "Deve ser informado um Segmento (ex: 09090900)";
                    return false;
                }
                return true;
            }

            if (tamanho_nivel == 6)
            {
                if (campo.Substring(4, 2) == "00" || campo.Substring(0, 2) == "00" || campo.Substring(2, 2) == "00")
                {
                    mensagem = "Deve ser informado um Segmento (ex: 090909)";
                    return false;
                }
                return true;
            }

            mensagem = $"Segmento não é válido para Subgrupo de {NivelCombo.Text} níveis.";
            return false;
        }

        private bool ValidarSubSegmento(string campo, int tamanho_nivel, ref string mensagem)
        {
            if (tamanho_nivel == 8)
            {
                if (campo.Substring(6, 2) == "00" || campo.Substring(0, 2) == "00" || campo.Substring(2, 2) == "00" || campo.Substring(4, 2) == "00")
                {
                    mensagem = "Deve ser informado um SubSegmento (ex: 09090909)";
                    return false;
                }
                return true;
            }

            mensagem = $"SubSegmento não é válido para Subgrupo de {NivelCombo.Text} níveis.";
            return false;
        }
    }
}