using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ValidarCSV
{
    public partial class Main : Form
    {
        public void Grid_limpar()
        {
            grid.DataSource = null;
            grid.Rows.Clear();
            grid.Columns.Clear();

            labellog.Text = "Registro:";

            Normal.Visible = false;
            Vermelho.Visible = false;
            Verde.Visible = false;
            Preto.Visible = false;
            Prata.Visible = false;
            Ouro.Visible = false;
            Mateador.Visible = false;
        }

        public void Grid_criar()
        {
            Grid_limpar();

            if (registros.Count == 0)
            {
                Tiao_definir();

                labellog.Text = "Nenhum erro encontrado";
                MessageBox.Show("Nenhum erro encontrado", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                labellog.Text = "Erros: " + registros.Count;

                DataTable TableGrid = new DataTable();

                TableGrid.Rows.Clear();

                grid.AllowUserToOrderColumns = true;
                grid.ReadOnly = true;

                TableGrid.Columns.Add("Campo", typeof(string));
                TableGrid.Columns.Add("Linha", typeof(string));
                TableGrid.Columns.Add("Coluna", typeof(string));
                TableGrid.Columns.Add("Valor", typeof(string));
                TableGrid.Columns.Add("Observacao", typeof(string));

                foreach (var registro in registros)
                {
                    DataRow row = TableGrid.NewRow();
                    row["Campo"] = registro.Campo;
                    row["Linha"] = registro.Linha;
                    row["Coluna"] = registro.Coluna;
                    row["Valor"] = registro.Valor;
                    row["Observacao"] = registro.Obs;
                    TableGrid.Rows.Add(row);
                }

                grid.DataSource = TableGrid;

                Zoom_grid_criar();
            }
        }

        private void Tiao_definir()
        {
            var aleatorio = new Random();

            // 20% de chance para exibir uma imagem
            bool exibirImagem = aleatorio.NextDouble() <= 0.20;

            string imagem = "";

            // Caso seja definido exibir uma imagem, aí sim verifica as porcentagens de cada tião
            if (exibirImagem)
            {
                var probabilidade = new Dictionary<string, double>
                {
                    { "Normal", 91.19 },
                    { "Vermelho", 5 },
                    { "Verde", 2 },
                    { "Preto", 1 },
                    { "Prata", 0.5 },
                    { "Ouro", 0.3 },
                    { "Mateador", 0.01 }
                };

                double totalChance = probabilidade.Values.Sum();
                double numeroAleatorio = aleatorio.NextDouble() * totalChance;

                double cumulativo = 0;
                foreach (var item in probabilidade)
                {
                    cumulativo += item.Value;
                    if (numeroAleatorio < cumulativo)
                    {
                        imagem = item.Key;
                        break;
                    }
                }
            }

            Normal.Visible = (imagem == "Normal");
            Vermelho.Visible = (imagem == "Vermelho");
            Verde.Visible = (imagem == "Verde");
            Preto.Visible = (imagem == "Preto");
            Prata.Visible = (imagem == "Prata");
            Ouro.Visible = (imagem == "Ouro");
            Mateador.Visible = (imagem == "Mateador");

        }

        public void Zoom_grid_limpar()
        {
            btnZoomIn.Visible = false;
            btnZoomOut.Visible = false;
            zoom.Visible = false;
        }

        private void Zoom_grid_criar()
        {
            //Zoom_grid_limpar();

            btnZoomIn.Visible = true;
            btnZoomOut.Visible = true;
            zoom.Visible = true;

            zoom.Text = "100%";
        }

        private void ZoomIn_click(object sender, EventArgs e)
        {
            Zoom_grid(grid, 2.0f);
        }

        private void ZoomOut_click(object sender, EventArgs e)
        {
            Zoom_grid(grid, -2.0f);
        }

        private void Zoom_grid(DataGridView dgv, float delta)
        {
            float currentFontSize = dgv.DefaultCellStyle.Font.Size;
            float newFontSize = currentFontSize + delta;

            if (newFontSize >= 6 && newFontSize <= 20)
            {
                dgv.DefaultCellStyle.Font = new System.Drawing.Font(dgv.DefaultCellStyle.Font.FontFamily, newFontSize);
                dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font(dgv.ColumnHeadersDefaultCellStyle.Font.FontFamily, newFontSize);

                if (newFontSize > currentFontSize)
                {
                    Zoom_label_atualizar(20);
                }
                else
                {
                    Zoom_label_atualizar(-20);
                }
            }
        }

        private void Zoom_label_atualizar(int increment)
        {
            int currentZoom = int.Parse(zoom.Text.Replace('%', ' ').Trim());
            currentZoom += increment;
            zoom.Text = currentZoom.ToString() + '%';
        }
    }
}