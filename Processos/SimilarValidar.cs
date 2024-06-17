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

    }
}