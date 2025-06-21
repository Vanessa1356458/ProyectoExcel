using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class BarraEstado
    {
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblPromedio;
        private ToolStripStatusLabel lblRecuento;
        private ToolStripStatusLabel lblSuma;
        private DataGridView dgvHoja;

        public StatusStrip StatusStrip => statusStrip;

        public BarraEstado(DataGridView dataGridView)
        {
            dgvHoja = dataGridView;
            CrearBarraEstado();
        }

        private void CrearBarraEstado()
        {
            statusStrip = new StatusStrip();
            statusStrip.BackColor = Color.FromArgb(240, 240, 240);

            lblPromedio = new ToolStripStatusLabel();
            lblPromedio.Text = "";
            lblPromedio.BorderSides = ToolStripStatusLabelBorderSides.Right;

            lblRecuento = new ToolStripStatusLabel();
            lblRecuento.Text = "";
            lblRecuento.BorderSides = ToolStripStatusLabelBorderSides.Right;

            lblSuma = new ToolStripStatusLabel();
            lblSuma.Text = "";

            statusStrip.Items.AddRange(new ToolStripItem[] { lblPromedio, lblRecuento, lblSuma });
        }

        public void Actualizar()
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                var valores = new List<double>();
                int recuento = 0;

                foreach (DataGridViewCell celda in dgvHoja.SelectedCells)
                {
                    string valorStr = celda.Value?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(valorStr) && double.TryParse(valorStr, out double valor))
                    {
                        valores.Add(valor);
                    }
                    if (!string.IsNullOrEmpty(valorStr))
                    {
                        recuento++;
                    }
                }

                if (valores.Count > 0)
                {
                    double suma = valores.Sum();
                    double promedio = valores.Average();

                    lblSuma.Text = $"Suma: {suma:N2}";
                    lblPromedio.Text = $"Promedio: {promedio:N2}";
                    lblRecuento.Text = $"Recuento: {recuento}";
                }
                else
                {
                    lblSuma.Text = "Suma: 0";
                    lblPromedio.Text = "Promedio: 0";
                    lblRecuento.Text = $"Recuento: {recuento}";
                }
            }
            else
            {
                LimpiarEstadisticas();
            }
        }

        public void LimpiarEstadisticas()
        {
            lblSuma.Text = "";
            lblPromedio.Text = "";
            lblRecuento.Text = "";
        }
    }
}
