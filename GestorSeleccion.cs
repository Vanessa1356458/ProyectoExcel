using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class GestorSeleccion
    {
        private DataGridView dgvHoja;
        private bool seleccionandoRango = false;
        private Point inicioSeleccion;
        private Point finSeleccion;
        private List<Point> celdasSeleccionadas = new List<Point>();
        private Color colorSeleccionPersonalizada = Color.FromArgb(100, 51, 153, 255);

        public bool SeleccionandoRango => seleccionandoRango;

        public GestorSeleccion(DataGridView dataGridView)
        {
            dgvHoja = dataGridView;
        }

        public void ManipularMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                var hit = dgvHoja.HitTest(e.X, e.Y);
                if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
                {
                    if (Control.ModifierKeys == Keys.Control)
                    {
                        Point celda = new Point(hit.ColumnIndex, hit.RowIndex);
                        if (celdasSeleccionadas.Contains(celda))
                            celdasSeleccionadas.Remove(celda);
                        else
                            celdasSeleccionadas.Add(celda);

                        ActualizarSeleccionVisual();
                    }
                    else if (Control.ModifierKeys == Keys.Shift)
                    {
                        if (dgvHoja.CurrentCell != null)
                        {
                            SeleccionarRango(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex,
                                           hit.ColumnIndex, hit.RowIndex);
                        }
                    }
                    else
                    {
                        celdasSeleccionadas.Clear();
                        inicioSeleccion = new Point(hit.ColumnIndex, hit.RowIndex);
                        seleccionandoRango = true;
                    }
                }
            }
        }

        public void ManipularMouseMove(MouseEventArgs e)
        {
            if (seleccionandoRango && e.Button == MouseButtons.Left)
            {
                var hit = dgvHoja.HitTest(e.X, e.Y);
                if (hit.RowIndex >= 0 && hit.ColumnIndex >= 0)
                {
                    finSeleccion = new Point(hit.ColumnIndex, hit.RowIndex);
                    SeleccionarRango(inicioSeleccion.X, inicioSeleccion.Y,
                                   finSeleccion.X, finSeleccion.Y);
                }
            }
        }

        public void ManipularMouseUp(MouseEventArgs e)
        {
            seleccionandoRango = false;
        }

        public void SeleccionarRango(int col1, int fila1, int col2, int fila2)
        {
            dgvHoja.ClearSelection();

            int minCol = Math.Min(col1, col2);
            int maxCol = Math.Max(col1, col2);
            int minFila = Math.Min(fila1, fila2);
            int maxFila = Math.Max(fila1, fila2);

            for (int col = minCol; col <= maxCol; col++)
            {
                for (int fila = minFila; fila <= maxFila; fila++)
                {
                    dgvHoja[col, fila].Selected = true;
                }
            }
        }

        private void ActualizarSeleccionVisual()
        {
            dgvHoja.ClearSelection();
            foreach (Point celda in celdasSeleccionadas)
            {
                dgvHoja[celda.X, celda.Y].Selected = true;
            }
        }

        public string ObtenerRangoSeleccionado()
        {
            if (dgvHoja.SelectedCells.Count == 0) return "";

            int minCol = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
            int maxCol = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);
            int minFila = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
            int maxFila = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);

            string inicioRango = $"{(char)('A' + minCol)}{minFila + 1}";
            string finRango = $"{(char)('A' + maxCol)}{maxFila + 1}";

            return minCol == maxCol && minFila == maxFila ? inicioRango : $"{inicioRango}:{finRango}";
        }

        public string ObtenerRangoSeleccionadoParaLabel()
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                var primeraCelda = dgvHoja.SelectedCells[dgvHoja.SelectedCells.Count - 1];
                var ultimaCelda = dgvHoja.SelectedCells[0];

                return $"{(char)('A' + primeraCelda.ColumnIndex)}{primeraCelda.RowIndex + 1}:" +
                       $"{(char)('A' + ultimaCelda.ColumnIndex)}{ultimaCelda.RowIndex + 1}";
            }
            return "";
        }

        public void InsertarFuncionFormula(TextBox txtFormula, string funcion)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = ObtenerRangoSeleccionado();
                txtFormula.Text = $"={funcion}({rango})";
                txtFormula.Focus();
            }
            else
            {
                txtFormula.Text = $"={funcion}(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }
    }
}
