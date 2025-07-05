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
        private Color colorResaltadoFormula = Color.FromArgb(150, 173, 216, 230);

        public bool SeleccionandoRango => seleccionandoRango;
        public List<Point> CeldasSeleccionadas => celdasSeleccionadas;

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
                        ManipularSeleccionControl(hit.ColumnIndex, hit.RowIndex);
                    }
                    else if (Control.ModifierKeys == Keys.Shift)
                    {
                        ManipularSeleccionShift(hit.ColumnIndex, hit.RowIndex);
                    }
                    else
                    {
                        IniciarSeleccionNormal(hit.ColumnIndex, hit.RowIndex);
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

        private void ManipularSeleccionControl(int col, int fila)
        {
            Point celda = new Point(col, fila);
            if (celdasSeleccionadas.Contains(celda))
                celdasSeleccionadas.Remove(celda);
            else
                celdasSeleccionadas.Add(celda);

            ActualizarSeleccionVisual();
        }

        private void ManipularSeleccionShift(int col, int fila)
        {
            if (dgvHoja.CurrentCell != null)
            {
                SeleccionarRango(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex,
                               col, fila);
            }
        }

        private void IniciarSeleccionNormal(int col, int fila)
        {
            celdasSeleccionadas.Clear();
            inicioSeleccion = new Point(col, fila);
            seleccionandoRango = true;
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
                    if (fila < dgvHoja.RowCount && col < dgvHoja.ColumnCount)
                    {
                        dgvHoja[col, fila].Selected = true;
                    }
                }
            }
        }

        private void ActualizarSeleccionVisual()
        {
            dgvHoja.ClearSelection();
            foreach (Point celda in celdasSeleccionadas)
            {
                if (celda.Y < dgvHoja.RowCount && celda.X < dgvHoja.ColumnCount)
                {
                    dgvHoja[celda.X, celda.Y].Selected = true;
                }
            }
        }

        public string ObtenerRangoSeleccionado()
        {
            if (dgvHoja.SelectedCells.Count == 0) return "";

            var celdas = dgvHoja.SelectedCells.Cast<DataGridViewCell>().ToList();

            int minCol = celdas.Min(c => c.ColumnIndex);
            int maxCol = celdas.Max(c => c.ColumnIndex);
            int minFila = celdas.Min(c => c.RowIndex);
            int maxFila = celdas.Max(c => c.RowIndex);

            string inicioRango = ConvertirACeldaRef(minCol, minFila);
            string finRango = ConvertirACeldaRef(maxCol, maxFila);

            return minCol == maxCol && minFila == maxFila ? inicioRango : $"{inicioRango}:{finRango}";
        }

        public string ObtenerRangoSeleccionadoParaLabel()
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                var celdas = dgvHoja.SelectedCells.Cast<DataGridViewCell>().ToList();
                var primeraCelda = celdas.OrderBy(c => c.RowIndex).ThenBy(c => c.ColumnIndex).First();
                var ultimaCelda = celdas.OrderByDescending(c => c.RowIndex).ThenByDescending(c => c.ColumnIndex).First();

                return $"{ConvertirACeldaRef(primeraCelda.ColumnIndex, primeraCelda.RowIndex)}:" +
                       $"{ConvertirACeldaRef(ultimaCelda.ColumnIndex, ultimaCelda.RowIndex)}";
            }
            else if (dgvHoja.CurrentCell != null)
            {
                return ConvertirACeldaRef(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex);
            }
            return "A1";
        }

        private string ConvertirACeldaRef(int columna, int fila)
        {
            return $"{(char)('A' + columna)}{fila + 1}";
        }

        public void InsertarFuncionFormula(TextBox txtFormula, string funcion)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = ObtenerRangoSeleccionado();
                txtFormula.Text = $"={funcion}({rango})";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
            else
            {
                txtFormula.Text = $"={funcion}(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }

        public void LimpiarSeleccion()
        {
            celdasSeleccionadas.Clear();
            dgvHoja.ClearSelection();
            seleccionandoRango = false;
        }

        public bool TieneSeleccionMultiple()
        {
            return dgvHoja.SelectedCells.Count > 1 || celdasSeleccionadas.Count > 1;
        }

        public List<DataGridViewCell> ObtenerCeldasSeleccionadas()
        {
            return dgvHoja.SelectedCells.Cast<DataGridViewCell>().ToList();
        }

        public void SeleccionarCelda(int columna, int fila)
        {
            if (fila >= 0 && fila < dgvHoja.RowCount && columna >= 0 && columna < dgvHoja.ColumnCount)
            {
                dgvHoja.ClearSelection();
                dgvHoja.CurrentCell = dgvHoja[columna, fila];
                dgvHoja[columna, fila].Selected = true;
            }
        }

        public void ExpandirSeleccion(int deltaCol, int deltaFila)
        {
            if (dgvHoja.CurrentCell != null)
            {
                int nuevaCol = Math.Max(0, Math.Min(dgvHoja.ColumnCount - 1, dgvHoja.CurrentCell.ColumnIndex + deltaCol));
                int nuevaFila = Math.Max(0, Math.Min(dgvHoja.RowCount - 1, dgvHoja.CurrentCell.RowIndex + deltaFila));

                if (Control.ModifierKeys == Keys.Shift)
                {
                    SeleccionarRango(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex, nuevaCol, nuevaFila);
                }
                else
                {
                    SeleccionarCelda(nuevaCol, nuevaFila);
                }
            }
        }
    }
}
