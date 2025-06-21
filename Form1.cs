using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Excel
{
    public partial class Form1 : Form
    {
        private DataGridView dgvHoja;
        private MenuStrip menuStrip;
        private ToolStrip toolStrip;
        private System.Windows.Forms.TextBox txtFormula;
        private Label lblCelda;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblPromedio;
        private ToolStripStatusLabel lblRecuento;
        private ToolStripStatusLabel lblSuma;
        private Dictionary<string, string> celdas;
        private string archivoActual = "";

        private bool seleccionandoRango = false;
        private Point inicioSeleccion;
        private Point finSeleccion;
        private List<Point> celdasSeleccionadas = new List<Point>();
        private Color colorSeleccionPersonalizada = Color.FromArgb(100, 51, 153, 255);

        public Form1()
        {
            InitializeComponentCustom();
            celdas = new Dictionary<string, string>();
        }
        public void SimularAbrirArchivo()
        {
            this.BeginInvoke(new Action(() => {
                System.Threading.Thread.Sleep(100);
                AbrirArchivo(null, null);
            }));
        }
        private void InitializeComponentCustom()
        {
            this.Text = "Excel Básico";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(240, 240, 240);

            CrearMenu();
            CrearToolbar();
            CrearAreaFormula();
            CrearDataGridView();
            CrearBarraEstado(); 
            ConfigurarEventos();
            ActualizarBarraEstado();
            dgvHoja.BringToFront();
        }
        private void CrearMenu()
        {
            menuStrip = new MenuStrip();
            menuStrip.BackColor = Color.FromArgb(250, 250, 250);
            menuStrip.Font = new Font("Segoe UI", 9F);

            var menuArchivo = new ToolStripMenuItem("Archivo");
            menuArchivo.DropDownItems.Add("Nuevo", null, NuevoArchivo);
            menuArchivo.DropDownItems.Add("Abrir", null, AbrirArchivo);
            menuArchivo.DropDownItems.Add("Guardar", null, GuardarArchivo);
            menuArchivo.DropDownItems.Add("Guardar Como", null, GuardarComo);
            menuArchivo.DropDownItems.Add(new ToolStripSeparator());
            menuArchivo.DropDownItems.Add("Salir", null, (s, e) => this.Close());

            var menuEdicion = new ToolStripMenuItem("Edición");
            menuEdicion.DropDownItems.Add("Copiar", null, Copiar);
            menuEdicion.DropDownItems.Add("Pegar", null, Pegar);
            menuEdicion.DropDownItems.Add("Cortar", null, Cortar);

            var menuInsertar = new ToolStripMenuItem("Insertar");
            menuInsertar.DropDownItems.Add("Fila", null, InsertarFila);
            menuInsertar.DropDownItems.Add("Columna", null, InsertarColumna);

            menuStrip.Items.AddRange(new ToolStripItem[] { menuArchivo, menuEdicion, menuInsertar });
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }
        private void CrearToolbar()
        {
            toolStrip = new ToolStrip();
            toolStrip.BackColor = Color.FromArgb(245, 245, 245);

            var btnNuevo = new ToolStripButton("Nuevo");
            btnNuevo.Click += NuevoArchivo;

            var btnAbrir = new ToolStripButton("Abrir");
            btnAbrir.Click += AbrirArchivo;

            var btnGuardar = new ToolStripButton("Guardar");
            btnGuardar.Click += GuardarArchivo;

            toolStrip.Items.AddRange(new ToolStripItem[] { btnNuevo, btnAbrir, btnGuardar });
            this.Controls.Add(toolStrip);
        }
        private void CrearAreaFormula()
        {
            var panel = new Panel();
            panel.Height = 35;
            panel.Dock = DockStyle.Top;
            panel.BackColor = Color.FromArgb(240, 240, 240);

            lblCelda = new Label();
            lblCelda.Text = "A1";
            lblCelda.Location = new Point(8, 8);
            lblCelda.Size = new Size(60, 22);
            lblCelda.BorderStyle = BorderStyle.FixedSingle;
            lblCelda.TextAlign = ContentAlignment.MiddleCenter;
            lblCelda.BackColor = Color.White;
            lblCelda.Font = new Font("Segoe UI", 9F);

            txtFormula = new System.Windows.Forms.TextBox();
            txtFormula.Location = new Point(75, 8);
            txtFormula.Size = new Size(400, 22);
            txtFormula.Font = new Font("Segoe UI", 9F);
            txtFormula.KeyDown += TxtFormula_KeyDown;

            panel.Controls.AddRange(new Control[] { lblCelda, txtFormula });
            this.Controls.Add(panel);
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
            this.Controls.Add(statusStrip);
        }
        private void CrearDataGridView()
        {
            dgvHoja = new DataGridView();

            dgvHoja.Dock = DockStyle.Fill;

            dgvHoja.Dock = DockStyle.Fill;
            dgvHoja.AllowUserToAddRows = false;
            dgvHoja.AllowUserToDeleteRows = false;
            dgvHoja.AllowUserToResizeRows = false;
            dgvHoja.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dgvHoja.ColumnHeadersHeight = 25;
            dgvHoja.RowHeadersWidth = 60;
            dgvHoja.GridColor = Color.FromArgb(208, 215, 229); 
            dgvHoja.BackgroundColor = Color.White;
            dgvHoja.BorderStyle = BorderStyle.Fixed3D;
            dgvHoja.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            dgvHoja.DefaultCellStyle.BackColor = Color.White;
            dgvHoja.DefaultCellStyle.ForeColor = Color.Black;
            dgvHoja.DefaultCellStyle.SelectionBackColor = Color.FromArgb(51, 153, 255);
            dgvHoja.DefaultCellStyle.SelectionForeColor = Color.White;
            dgvHoja.DefaultCellStyle.Font = new Font("Calibri", 11F);

            dgvHoja.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(198, 217, 191);
            dgvHoja.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvHoja.ColumnHeadersDefaultCellStyle.Font = new Font("Calibri", 11F, FontStyle.Bold);
            dgvHoja.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvHoja.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(198, 217, 191);
            dgvHoja.RowHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvHoja.RowHeadersDefaultCellStyle.Font = new Font("Calibri", 11F, FontStyle.Bold);
            dgvHoja.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 0; i < 26; i++)
            {
                var columna = new DataGridViewTextBoxColumn();
                columna.Name = ((char)('A' + i)).ToString();
                columna.HeaderText = columna.Name;
                columna.Width = 80;
                columna.SortMode = DataGridViewColumnSortMode.NotSortable;
                dgvHoja.Columns.Add(columna);
            }

            string[] filasVacias = new string[26];
            for (int i = 0; i < 26; i++)
            {
                filasVacias[i] = "";
            }

            for (int i = 0; i < 100; i++)
            {
                int index = dgvHoja.Rows.Add(filasVacias);
                dgvHoja.Rows[index].HeaderCell.Value = (i + 1).ToString();
                dgvHoja.Rows[index].Height = 20;
            }

            dgvHoja.ResumeLayout();
            this.Controls.Add(dgvHoja);
        }
        private void ConfigurarEventos()
        {
            dgvHoja.CellClick += DgvHoja_CellClick;
            dgvHoja.CellEndEdit += DgvHoja_CellEndEdit;
            dgvHoja.CellEnter += DgvHoja_CellEnter;
            dgvHoja.EditMode = DataGridViewEditMode.EditOnKeystroke;
            dgvHoja.SelectionChanged += DgvHoja_SelectionChanged;
            dgvHoja.MouseDown += DgvHoja_MouseDown;
            dgvHoja.MouseMove += DgvHoja_MouseMove;
            dgvHoja.MouseUp += DgvHoja_MouseUp;
            dgvHoja.KeyDown += DgvHoja_KeyDown; // AÑADIR este evento que faltaba

            dgvHoja.StandardTab = true;
            dgvHoja.AllowUserToResizeColumns = true;
            dgvHoja.AllowUserToResizeRows = true;

            dgvHoja.RowHeadersVisible = true;
            dgvHoja.ColumnHeadersVisible = true;
            dgvHoja.RowHeadersWidth = 50;
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (dgvHoja != null)
            {
                dgvHoja.Width = this.ClientSize.Width;
                dgvHoja.Height = this.ClientSize.Height - 140;
            }
        }
        private void ActualizarBarraEstado()
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
                lblSuma.Text = "";
                lblPromedio.Text = "";
                lblRecuento.Text = "";
            }
        }
        private void DgvHoja_MouseDown(object sender, MouseEventArgs e)
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
        private void DgvHoja_MouseMove(object sender, MouseEventArgs e)
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
        private void DgvHoja_MouseUp(object sender, MouseEventArgs e)
        {
            seleccionandoRango = false;
        }
        private void DgvHoja_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                var primeraCelda = dgvHoja.SelectedCells[dgvHoja.SelectedCells.Count - 1];
                var ultimaCelda = dgvHoja.SelectedCells[0];

                string rango = $"{(char)('A' + primeraCelda.ColumnIndex)}{primeraCelda.RowIndex + 1}:" +
                              $"{(char)('A' + ultimaCelda.ColumnIndex)}{ultimaCelda.RowIndex + 1}";
                lblCelda.Text = rango;
            }

            ActualizarBarraEstado();
        }
        private void SeleccionarRango(int col1, int fila1, int col2, int fila2)
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
        private void InsertarSuma(object sender, EventArgs e)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = ObtenerRangoSeleccionado();
                txtFormula.Text = $"=SUM({rango})";
                txtFormula.Focus();
            }
            else
            {
                txtFormula.Text = "=SUM(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }
        private void InsertarPromedio(object sender, EventArgs e)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = ObtenerRangoSeleccionado();
                txtFormula.Text = $"=AVERAGE({rango})";
                txtFormula.Focus();
            }
            else
            {
                txtFormula.Text = "=AVERAGE(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }
        private void InsertarContar(object sender, EventArgs e)
        {
            if (dgvHoja.SelectedCells.Count > 1)
            {
                string rango = ObtenerRangoSeleccionado();
                txtFormula.Text = $"=COUNT({rango})";
                txtFormula.Focus();
            }
            else
            {
                txtFormula.Text = "=COUNT(";
                txtFormula.Focus();
                txtFormula.SelectionStart = txtFormula.Text.Length;
            }
        }
        private string ObtenerRangoSeleccionado()
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
        private void DgvHoja_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            var celda = dgvHoja[e.ColumnIndex, e.RowIndex];

            if (celda.Tag is string formula && formula.StartsWith("="))
            {
                celda.Value = formula;
            }
        }
        private void DgvHoja_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                if (dgvHoja.CurrentCell != null && !dgvHoja.IsCurrentCellInEditMode)
                {
                    var celda = dgvHoja.CurrentCell;
                    string celdaRef = $"{(char)('A' + celda.ColumnIndex)}{celda.RowIndex + 1}";

                    celda.Value = "";
                    celda.Tag = null;
                    celda.Style.ForeColor = Color.Black;

                    if (celdas.ContainsKey(celdaRef))
                        celdas.Remove(celdaRef);

                    txtFormula.Text = "";

                    RecalcularFormulas();

                    e.Handled = true;
                }
            }
        }
        private void DgvHoja_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                string celda = $"{(char)('A' + e.ColumnIndex)}{e.RowIndex + 1}";
                lblCelda.Text = celda;

                if (celdas.TryGetValue(celda, out string valor))
                {
                    txtFormula.Text = valor;
                }
                else if (dgvHoja[e.ColumnIndex, e.RowIndex].Tag is string formula)
                {
                    txtFormula.Text = formula;
                }
                else
                {
                    txtFormula.Text = dgvHoja[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
                }   
            }
        }
        private void DgvHoja_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            DgvHoja_CellClick(sender, e);
        }
        private void DgvHoja_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            string celda = $"{(char)('A' + e.ColumnIndex)}{e.RowIndex + 1}";
            string valor = dgvHoja[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";

            celdas[celda] = valor;

            if (valor.StartsWith("="))
            {
                double resultado = Formulas.Evaluar(valor, dgvHoja);

                dgvHoja[e.ColumnIndex, e.RowIndex].Tag = valor;

                if (double.IsNaN(resultado))
                {
                    dgvHoja[e.ColumnIndex, e.RowIndex].Value = "#ERROR";
                    dgvHoja[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.Red;
                }
                else
                {
                    dgvHoja[e.ColumnIndex, e.RowIndex].Value = resultado;
                    dgvHoja[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.Black;
                }
            }
            else
            {
                dgvHoja[e.ColumnIndex, e.RowIndex].Tag = null;
                dgvHoja[e.ColumnIndex, e.RowIndex].Style.ForeColor = Color.Black;
            }

            RecalcularFormulas();
            ActualizarBarraEstado();
        }
        private void RecalcularFormulas()
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
        private void TxtFormula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && dgvHoja.CurrentCell != null)
            {
                string valor = txtFormula.Text;
                int fila = dgvHoja.CurrentCell.RowIndex;
                int columna = dgvHoja.CurrentCell.ColumnIndex;
                string celda = $"{(char)('A' + columna)}{fila + 1}";

                celdas[celda] = valor;

                if (valor.StartsWith("="))
                {
                    double resultado = Formulas.Evaluar(valor, dgvHoja);
                    dgvHoja[columna, fila].Tag = valor; 

                    if (double.IsNaN(resultado))
                    {
                        dgvHoja[columna, fila].Value = "#ERROR";
                        dgvHoja[columna, fila].Style.ForeColor = Color.Red;
                    }
                    else
                    {
                        dgvHoja[columna, fila].Value = resultado;
                        dgvHoja[columna, fila].Style.ForeColor = Color.Black;
                    }
                }
                else
                {
                    if (double.TryParse(valor, out double num))
                    {
                        dgvHoja[columna, fila].Value = num;
                    }
                    else
                    {
                        dgvHoja[columna, fila].Value = valor;
                    }
                    dgvHoja[columna, fila].Tag = null; 
                    dgvHoja[columna, fila].Style.ForeColor = Color.Black;
                }

                RecalcularFormulas();
            }
        }
        private void NuevoArchivo(object sender, EventArgs e)
        {
            dgvHoja.Rows.Clear();
            celdas.Clear();

            for (int i = 0; i < 100; i++)
            {
                int index = dgvHoja.Rows.Add();
                dgvHoja.Rows[index].HeaderCell.Value = (i + 1).ToString();
            }

            archivoActual = "";
            this.Text = "Excel Básico - Nuevo archivo";
        }
        private void AbrirArchivo(object sender, EventArgs e)
        {
            using (var openDialog = new OpenFileDialog())
            {
                openDialog.Filter = "Archivos CSV|*.csv|Todos los archivos|*.*";

                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        NuevoArchivo(sender, e);
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
                        this.Text = $"Excel Básico - {Path.GetFileName(archivoActual)}";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al abrir archivo: {ex.Message}");
                    }
                }
            }
        }
        private void GuardarArchivo(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(archivoActual))
                GuardarComo(sender, e);
            else
                GuardarEnArchivo(archivoActual);
        }
        private void GuardarComo(object sender, EventArgs e)
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Archivos CSV|*.csv|Todos los archivos|*.*";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    archivoActual = saveDialog.FileName;
                    GuardarEnArchivo(archivoActual);
                    this.Text = $"Excel Básico - {Path.GetFileName(archivoActual)}";
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
        private void Copiar(object sender, EventArgs e)
        {
            if (dgvHoja.CurrentCell != null)
                Clipboard.SetText(dgvHoja.CurrentCell.Value?.ToString() ?? "");
        }
        private void Pegar(object sender, EventArgs e)
        {
            if (dgvHoja.CurrentCell != null && Clipboard.ContainsText())
                dgvHoja.CurrentCell.Value = Clipboard.GetText();
        }
        private void Cortar(object sender, EventArgs e)
        {
            if (dgvHoja.CurrentCell != null)
            {
                Clipboard.SetText(dgvHoja.CurrentCell.Value?.ToString() ?? "");
                dgvHoja.CurrentCell.Value = "";
            }
        }
        private void InsertarFila(object sender, EventArgs e)
        {
            if (dgvHoja.CurrentCell != null)
            {
                int indice = dgvHoja.CurrentCell.RowIndex;
                dgvHoja.Rows.Insert(indice, 1);

                for (int i = 0; i < dgvHoja.Rows.Count; i++)
                    dgvHoja.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }
        }
        private void InsertarColumna(object sender, EventArgs e)
        {
            if (dgvHoja.CurrentCell != null && dgvHoja.Columns.Count < 26)
            {
                int indice = dgvHoja.CurrentCell.ColumnIndex;
                dgvHoja.Columns.Insert(indice, new DataGridViewTextBoxColumn
                {
                    Name = ((char)('A' + dgvHoja.Columns.Count)).ToString(),
                    HeaderText = ((char)('A' + dgvHoja.Columns.Count)).ToString(),
                    Width = 80
                });
            }
        }
    }
    
}
        
    

