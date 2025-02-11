﻿using ClosedXML.Excel;
using MathNet.Numerics;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using static ValidarCSV.TypeExtensions;
using System.Linq;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        public void Campos_validar_gerenciar(string tabela, string campo, int linha, int coluna, TipoCampoType tipo, double tamanho_formato, Boolean obrigatorio)
        {
            string mensagem = string.Empty;
            bool valido = true;

            if (Obrigatorio_validar(campo, tipo, obrigatorio, tamanho_formato, ref mensagem))
            {
                Registro_adicionar(tabela, linha, coluna, campo, mensagem);
                return;
            }

            switch (tipo)
            {
                //campos padrão
                case TipoCampoType.Character:
                    Char_validar(campo, tamanho_formato, ref mensagem, ref valido);
                    break;

                case TipoCampoType.Numeric:
                    Numeric_validar(campo.Trim(), tamanho_formato, ref mensagem, ref valido);
                    break;

                case TipoCampoType.Integer:
                    Integer_validar(campo, tamanho_formato, ref mensagem, ref valido);
                    break;

                case TipoCampoType.Date:
                    Date_validar(campo.Trim(), ref mensagem, ref valido);
                    break;

                //Campos 'especiais'
                case TipoCampoType.DateFormat:
                    Date_formato_validar(campo.Trim(), Formato_date_retornar(tamanho_formato), ref mensagem, ref valido);
                    break;

                case TipoCampoType.Nivel: //Grupos e Subgrupos
                    Nivel_validar(campo.Trim(), tamanho_formato, ref mensagem, ref valido);
                    break;

                case TipoCampoType.Dominio: //provindos de enum do genexus
                    Dominio_validar(campo, tamanho_formato, obrigatorio, ref mensagem, ref valido);
                    break;

                case TipoCampoType.Sintetica:
                    Sintetica_validar(campo, tamanho_formato, obrigatorio, ref mensagem, ref valido);
                    break;

                case TipoCampoType.InscricaoEstadual:
                    Insricao_estadual_validar(campo, tamanho_formato, obrigatorio, ref mensagem, ref valido);
                    break;

                default:
                    Registro_adicionar(tabela, linha, coluna, campo, "Validação falhou, conferir manualmente");
                    break;
            }

            if (!valido)
            {
                Registro_adicionar(tabela, linha, coluna, campo, mensagem);
            }
        }

        public bool Obrigatorio_validar(string campo, TipoCampoType tipo, bool obrigatorio, double tamanho_formato, ref string mensagem)
        {
            if (!obrigatorio)
            {
                return false;
            }

            mensagem = string.Empty;

            if (tipo == TipoCampoType.Integer || tipo == TipoCampoType.Numeric || tipo == TipoCampoType.Nivel)
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

            List<string> invalidos = Dominio_lista_retornar(Dominio_retornar(DominioType.Invalidos));

            if (tipo == TipoCampoType.Dominio)
            {
                List<string> dominio = Dominio_lista_retornar(tamanho_formato);
                List<string> dominioResultado = invalidos.Except(dominio).ToList();

                if (dominioResultado.Contains(campo.Trim()))
                {
                    mensagem = "Campo obrigatório";
                }
            }
            else
            {
                if (string.IsNullOrEmpty(mensagem) && (invalidos.Contains(campo.Trim())))
                {
                    mensagem = "Campo obrigatório";
                }
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

        private void Char_validar(string campo, double tamanho_formato, ref string mensagem, ref bool valido)
        {
            if (campo.Length > tamanho_formato)
            {
                valido = false;
                mensagem = "Excede " + tamanho_formato.ToString() + " caracteres";
            }
        }

        private void Numeric_validar(string valor, double tamanho_formato, ref string mensagem_erro, ref bool valido)
        {
            valor = valor.Replace(".", "");
            if (valor == "0" || valor.Trim() == "")
            {
                return;
            }

            int precisao = (int)Math.Truncate(tamanho_formato);
            double parteDecimal = (tamanho_formato - precisao).Round(1);
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
                mensagem = "Deve estar em um formato de data válido, conforme layout: " + formato.ToUpper().Trim();
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
            campo = campo.ToUpper();
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

        private void Sintetica_validar(string campo, double tamanho_formato, bool obrigatorio, ref string mensagem, ref bool valido)
        {
            if (campo.StartsWith(".") || campo.EndsWith("."))
            {
                mensagem = "Não pode haver ponto no início ou no final";
            }

            // Verificar dois pontos seguidos
            if (campo.Contains(".."))
            {
                Adicionar_mensagem("Não pode haver dois pontos seguidos", ref mensagem);
            }

            // Dividir os níveis e validar individualmente
            string[] levels = campo.Split('.');

            if (levels.Length > 7)
            {
                Adicionar_mensagem("Mais de 7 níveis", ref mensagem);
            }

            for (int i = 0; i < levels.Length; i++)
            {
                // Validar o 1º, 6º e 7º nível (1 dígito apenas)
                if ((i == 0 || i == 5 || i == 6) && !Regex.IsMatch(levels[i], @"^\d{1}$"))
                {
                    Adicionar_mensagem($"O nível {i + 1} deve conter exatamente 1 dígito", ref mensagem);
                }

                // Validar o 2º ao 5º nível (exatamente 2 dígitos)
                if ((i >= 1 && i <= 4) && !Regex.IsMatch(levels[i], @"^\d{2}$"))
                {
                    Adicionar_mensagem($"O nível {i + 1} deve conter exatamente 2 dígitos", ref mensagem);
                }
            }

            if (!string.IsNullOrEmpty(mensagem))
            {
                valido = false;
                return;
            }
        }

        private void Insricao_estadual_validar(string campo, double tamanho_formato, bool obrigatorio, ref string mensagem, ref bool valido)
        {
            if (campo == "0" || campo.Trim() == "" /*|| campo.Trim().ToUpper().Equals("ISENTO") || campo.Trim().ToUpper().Equals("ISENTA")*/ )
            {
                return;
            }

            int Tamanho1 = (int)Math.Truncate(tamanho_formato);
            double parteDecimal = (tamanho_formato - Tamanho1).Round(1);
            int Tamanho2 = (int)(parteDecimal * 10);

            bool isInteiro = !long.TryParse(campo.Trim(), out _);

            if (
               isInteiro                                                               ||
               (Tamanho2 > 0 && campo.Length != Tamanho1 && campo.Length != Tamanho2)  ||
               (Tamanho1 < 20 && Tamanho2 == 0 && campo.Length != Tamanho1)            ||
               (Tamanho1 == 20 && Tamanho2 == 0 && campo.Length > Tamanho1)
               )
            {
                valido = false;

                string mensagem_inteiro = isInteiro ? " ser um número inteiro e" : "";
                if (Tamanho2 == 0)
                {
                    mensagem = "Deve" + mensagem_inteiro + " conter até " + Tamanho1 + " dígitos";
                }
                else
                {
                    mensagem = "Deve" + mensagem_inteiro + " conter " + Tamanho1 + " ou " + Tamanho2 + " dígitos";
                }
            }

        }

        private void Adicionar_mensagem(string adicionar, ref string mensagem)
        {
            if (string.IsNullOrWhiteSpace(mensagem))
            {
                mensagem = adicionar;
            }
            else
            {
                mensagem += ". " + adicionar;
            }
        }

        private void Sobressalente_validar(int rows, int columns, string campo) //Chamado diretamente no layout caso as colunas ultrapassem o cabeçalho
        {
            List<string> invalidos = Dominio_lista_retornar(Dominio_retornar(DominioType.Invalidos));
            if (!string.IsNullOrEmpty(campo) || !invalidos.Contains(campo.Trim()))
            {
                Registro_adicionar("Erro genérico", rows, columns, campo, "Excedeu o número de colunas do cabeçalho");
            }
        }
    }
}