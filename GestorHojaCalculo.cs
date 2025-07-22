using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class GestorHojaCalculo
    {
        private DataGridView dgvHoja;
        private TextBox txtFormula;
        private Label lblCelda;
        private Dictionary<string, string> celdas;
        private GestorSeleccion gestorSeleccion;

        public GestorHojaCalculo(DataGridView dataGridView, TextBox textBoxFormula, Label labelCelda, GestorSeleccion gestorSel)
        {
            dgvHoja = dataGridView;
            txtFormula = textBoxFormula;
            lblCelda = labelCelda;
            gestorSeleccion = gestorSel;
            celdas = new Dictionary<string, string>();

            txtFormula.KeyDown += TxtFormula_KeyDown;
            dgvHoja.CellClick += DgvHoja_CellClick;
        }

        private void TxtFormula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && dgvHoja.CurrentCell != null)
            {
                ProcesarEntradaCelda();
            }
        }

        private void DgvHoja_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                ActualizarFormulaBar(e.ColumnIndex, e.RowIndex);
            }
        }

        private void ProcesarEntradaCelda()
        {
            if (dgvHoja.CurrentCell == null) return;

            string valor = txtFormula.Text;
            int fila = dgvHoja.CurrentCell.RowIndex;
            int columna = dgvHoja.CurrentCell.ColumnIndex;
            string celdaKey = ObtenerNombreCelda(columna, fila);
            
            celdas[celdaKey] = valor;

            if (valor.StartsWith("="))
            {
                ProcesarFormula(columna, fila, valor);
            }
            else
            {
                ProcesarValorDirecto(columna, fila, valor);
            }

            RecalcularTodasLasFormulas();
        }

        private void ActualizarFormulaBar(int columna, int fila)
        {
            string celda = ObtenerNombreCelda(columna, fila);
            lblCelda.Text = celda;

            if (celdas.TryGetValue(celda, out string valor))
            {
                txtFormula.Text = valor;
            }
            else if (dgvHoja[columna, fila].Tag is string formula)
            {
                txtFormula.Text = formula;
            }
            else
            {
                txtFormula.Text = dgvHoja[columna, fila].Value?.ToString() ?? "";
            }
        }

        private void ProcesarFormula(int columna, int fila, string formula)
        {
            try
            {
                double resultado = Formulas.Evaluar(formula, dgvHoja);
                var celda = dgvHoja[columna, fila];

                celda.Tag = formula; 

                if (double.IsNaN(resultado))
                {
                    celda.Value = "#ERROR";
                    celda.Style.ForeColor = Color.Red;
                }
                else
                {
                    celda.Value = resultado;
                    celda.Style.ForeColor = Color.Black;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error procesando fórmula: {ex.Message}");
                dgvHoja[columna, fila].Value = "#ERROR";
                dgvHoja[columna, fila].Style.ForeColor = Color.Red;
            }
        }

        private void ProcesarValorDirecto(int columna, int fila, string valor)
        {
            var celda = dgvHoja[columna, fila];

            if (double.TryParse(valor, out double num))
            {
                celda.Value = num;
            }
            else
            {
                celda.Value = valor;
            }

            celda.Tag = null; 
            celda.Style.ForeColor = Color.Black;
        }

        private void RecalcularTodasLasFormulas()
        {
            try
            {
                for (int row = 0; row < dgvHoja.Rows.Count; row++)
                {
                    for (int col = 0; col < dgvHoja.Columns.Count; col++)
                    {
                        var celda = dgvHoja[col, row];
                        if (celda.Tag is string formula && formula.StartsWith("="))
                        {
                            double resultado = Formulas.Evaluar(formula, dgvHoja);

                            if (double.IsNaN(resultado))
                            {
                                celda.Value = "#ERROR";
                                celda.Style.ForeColor = Color.Red;
                            }
                            else
                            {
                                celda.Value = resultado;
                                celda.Style.ForeColor = Color.Black;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error recalculando fórmulas: {ex.Message}");
            }
        }

        public void InsertarFormula(string tipoFormula)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = gestorSeleccion.ObtenerRangoSeleccionado();
                txtFormula.Text = $"={tipoFormula}({rango})";
                txtFormula.Focus();
            }
            else
            {
                txtFormula.Text = $"={tipoFormula}(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }

        public void LimpiarCelda(int columna, int fila)
        {
            string nombreCelda = ObtenerNombreCelda(columna, fila);

            if (celdas.ContainsKey(nombreCelda))
            {
                celdas.Remove(nombreCelda);
            }

            var celda = dgvHoja[columna, fila];
            celda.Value = null;
            celda.Tag = null;
            celda.Style.ForeColor = Color.Black;

            RecalcularTodasLasFormulas();
        }

        public void LimpiarCeldaActual()
        {
            if (dgvHoja.CurrentCell != null)
            {
                LimpiarCelda(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex);
                ActualizarFormulaBar(dgvHoja.CurrentCell.ColumnIndex, dgvHoja.CurrentCell.RowIndex);
            }
        }

        public string ObtenerValorOriginalCelda(int columna, int fila)
        {
            string nombreCelda = ObtenerNombreCelda(columna, fila);
            return celdas.TryGetValue(nombreCelda, out string valor) ? valor : "";
        }

        public bool CeldaTieneFormula(int columna, int fila)
        {
            return dgvHoja[columna, fila].Tag is string formula && formula.StartsWith("=");
        }

        private string ObtenerNombreCelda(int columna, int fila)
        {
            return $"{(char)('A' + columna)}{fila + 1}";
        }

        public Dictionary<string, string> ObtenerTodasLasCeldas()
        {
            return new Dictionary<string, string>(celdas);
        }

        public void CargarCeldas(Dictionary<string, string> celdasGuardadas)
        {
            celdas.Clear();

            foreach (var kvp in celdasGuardadas)
            {
                celdas[kvp.Key] = kvp.Value;

                if (ParsearReferenciaCelda(kvp.Key, out int col, out int fila))
                {
                    if (kvp.Value.StartsWith("="))
                    {
                        ProcesarFormula(col, fila, kvp.Value);
                    }
                    else
                    {
                        ProcesarValorDirecto(col, fila, kvp.Value);
                    }
                }
            }

            RecalcularTodasLasFormulas();
        }

        private bool ParsearReferenciaCelda(string referencia, out int columna, out int fila)
        {
            columna = -1;
            fila = -1;

            if (string.IsNullOrWhiteSpace(referencia) || referencia.Length < 2)
                return false;

            try
            {
                int i = 0;
                while (i < referencia.Length && char.IsLetter(referencia[i]))
                {
                    i++;
                }

                if (i == 0 || i >= referencia.Length)
                    return false;

                string columnaStr = referencia.Substring(0, i);
                string filaStr = referencia.Substring(i);

                columna = 0;
                for (int j = 0; j < columnaStr.Length; j++)
                {
                    columna = columna * 26 + (columnaStr[j] - 'A' + 1);
                }
                columna--; 

                if (!int.TryParse(filaStr, out fila) || fila <= 0)
                    return false;

                fila--; 

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
