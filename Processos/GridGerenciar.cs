using System;
using System.Data;
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
        }

        public void Grid_criar()
        {
            Grid_limpar();

            if (registros.Count == 0)
            {
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

        public void Zoom_grid_limpar()
        {
            btnZoomIn.Visible = false;
            btnZoomOut.Visible = false;
            zoom.Visible = false;
        }

        private void Zoom_grid_criar()
        {
            Zoom_grid_limpar();

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