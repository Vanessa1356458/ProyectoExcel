using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class GestorArchivos
    {
        private DataGridView dgvHoja;
        private Dictionary<string, string> celdas;
        private string archivoActual = "";
        private Form formularioPrincipal;
        public string ArchivoActual
        {
            get => archivoActual;
            private set => archivoActual = value;
        }
        public GestorArchivos(DataGridView dataGridView, Dictionary<string, string> celdasDictionary, Form form)
        {
            dgvHoja = dataGridView;
            celdas = celdasDictionary;
            formularioPrincipal = form;
        }
        public void NuevoArchivo()
        {
            dgvHoja.Rows.Clear();
            celdas.Clear();

            for (int i = 0; i < 100; i++)
            {
                int index = dgvHoja.Rows.Add();
                dgvHoja.Rows[index].HeaderCell.Value = (i + 1).ToString();
            }

            archivoActual = "";
            formularioPrincipal.Text = "Excel Básico - Nuevo archivo";
        }
        public void AbrirArchivo()
        {
            using (var openDialog = new OpenFileDialog())
            {
                openDialog.Filter = "Archivos CSV|*.csv|Todos los archivos|*.*";

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        NuevoArchivo();
                        string[] lineas = File.ReadAllLines(openDialog.FileName);

                        for (int fila = 0; fila < lineas.Length && fila < dgvHoja.Rows.Count; fila++)
                        {
                            string[] valores = lineas[fila].Split(',');
                            for (int col = 0; col < valores.Length && col < dgvHoja.Columns.Count; col++)
                            {
                                dgvHoja[col, fila].Value = valores[col];
                            }
                        }

                        archivoActual = openDialog.FileName;
                        formularioPrincipal.Text = $"Excel Básico - {Path.GetFileName(archivoActual)}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al abrir archivo: {ex.Message}");
                    }
                }
            }
        }
        public void GuardarArchivo()
        {
            if (string.IsNullOrEmpty(archivoActual))
                GuardarComo();
            else
                GuardarEnArchivo(archivoActual);
        }
        public void GuardarComo()
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Archivos CSV|*.csv|Todos los archivos|*.*";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    archivoActual = saveDialog.FileName;
                    GuardarEnArchivo(archivoActual);
                    formularioPrincipal.Text = $"Excel Básico - {Path.GetFileName(archivoActual)}";
                }
            }
        }
        private void GuardarEnArchivo(string archivo)
        {
            try
            {
                var sb = new StringBuilder();

                for (int fila = 0; fila < dgvHoja.Rows.Count; fila++)
                {
                    var valores = new List<string>();
                    bool filaVacia = true;

                    for (int col = 0; col < dgvHoja.Columns.Count; col++)
                    {
                        string valor = dgvHoja[col, fila].Value?.ToString() ?? "";
                        valores.Add(valor);
                        if (!string.IsNullOrEmpty(valor)) filaVacia = false;
                    }

                    if (!filaVacia)
                        sb.AppendLine(string.Join(",", valores));
                }

                File.WriteAllText(archivo, sb.ToString());
                MessageBox.Show("Archivo guardado exitosamente");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}");
            }
        }
        public void Copiar()
        {
            if (dgvHoja.SelectedCells.Count > 0)
            {
                var sb = new StringBuilder();

                int minFila = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
                int maxFila = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);
                int minCol = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
                int maxCol = dgvHoja.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);

                for (int fila = minFila; fila <= maxFila; fila++)
                {
                    for (int col = minCol; col <= maxCol; col++)
                    {
                        if (col > minCol) sb.Append('\t');

                        var celda = dgvHoja[col, fila];
                        sb.Append(celda.Value?.ToString() ?? "");
                    }
                    sb.AppendLine();
                }

                Clipboard.SetText(sb.ToString());
            }
        }
        public void Pegar()
        {
            if (dgvHoja.CurrentCell != null && Clipboard.ContainsText())
            {
                string texto = Clipboard.GetText();

                string[] lineas = texto.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                int filaInicial = dgvHoja.CurrentCell.RowIndex;
                int colInicial = dgvHoja.CurrentCell.ColumnIndex;

                for (int i = 0; i < lineas.Length; i++)
                {
                    if (i == lineas.Length - 1 && string.IsNullOrEmpty(lineas[i]))
                        continue;

                    string[] columnas = lineas[i].Split('\t');

                    for (int j = 0; j < columnas.Length; j++)
                    {
                        int filaDestino = filaInicial + i;
                        int colDestino = colInicial + j;

                        if (filaDestino < dgvHoja.RowCount && colDestino < dgvHoja.ColumnCount)
                        {
                            try
                            {
                                dgvHoja[colDestino, filaDestino].Value = columnas[j];
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error al pegar en celda [{filaDestino}, {colDestino}]: {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
        public void Cortar()
        {
            if (dgvHoja.SelectedCells.Count > 0)
            {
                Copiar();

                foreach (DataGridViewCell celda in dgvHoja.SelectedCells)
                {
                    try
                    {
                        celda.Value = "";
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error al cortar celda [{celda.RowIndex}, {celda.ColumnIndex}]: {ex.Message}");
                    }
                }
            }
            else if (dgvHoja.CurrentCell != null)
            {
                Clipboard.SetText(dgvHoja.CurrentCell.Value?.ToString() ?? "");
                dgvHoja.CurrentCell.Value = "";
            }
        }
        public void InsertarFila()
        {
            if (dgvHoja.CurrentCell != null)
            {
                int indice = dgvHoja.CurrentCell.RowIndex;
                dgvHoja.Rows.Insert(indice, 1);

                for (int i = 0; i < dgvHoja.Rows.Count; i++)
                    dgvHoja.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }
        public void InsertarColumna()
        {
            if (dgvHoja.CurrentCell != null)
            {
                int indice = dgvHoja.CurrentCell.ColumnIndex;

                var nuevaColumna = new DataGridViewTextBoxColumn
                {
                    Name = "temp",
                    HeaderText = "temp",
                    Width = 80,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };

                dgvHoja.Columns.Insert(indice, nuevaColumna);

                for (int i = 0; i < dgvHoja.Columns.Count; i++)
                {
                    int n = i;
                    string nombreColumna = "";

                    while (n >= 0)
                    {
                        nombreColumna = (char)('A' + (n % 26)) + nombreColumna;
                        n = (n / 26) - 1;
                    }

                    dgvHoja.Columns[i].Name = nombreColumna;
                    dgvHoja.Columns[i].HeaderText = nombreColumna;
                }

                for (int i = 0; i < dgvHoja.Rows.Count; i++)
                {
                    dgvHoja.Rows[i].Cells[indice].Value = "";
                }
            }
            else
            {
                MessageBox.Show("Selecciona una celda para insertar una columna.");
            }
        }
    }
}
