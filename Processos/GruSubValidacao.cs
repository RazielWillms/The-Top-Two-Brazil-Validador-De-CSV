
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        private bool Validar_Grupo(string campo, int tamanho_nivel, ref string mensagem)
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

        private bool Validar_SubNivel(string campo, int tamanho_nivel, ref string mensagem, string nivel)
        {
            switch (nivel)
            {
                case "SubGrupo":
                    return Validar_SubGrupo(campo, tamanho_nivel, ref mensagem);

                case "Segmento":
                    return Validar_Segmento(campo, tamanho_nivel, ref mensagem);

                case "SubSegmento":
                    return Validar_SubSegmento(campo, tamanho_nivel, ref mensagem);

                default:
                    mensagem = "Nível desconhecido";
                    return false;
            }
        }

        private bool Validar_SubGrupo(string campo, int tamanho_nivel, ref string mensagem)
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

        private bool Validar_Segmento(string campo, int tamanho_nivel, ref string mensagem)
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

        private bool Validar_SubSegmento(string campo, int tamanho_nivel, ref string mensagem)
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