using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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
        private Dictionary<string, string> celdas;
        private string archivoActual = "";
        private GestorHojaCalculo gestorHoja;
        private GestorArchivos gestorArchivos;
        private GestorSeleccion gestorSeleccion;
        private BarraEstado barraEstado;
        private bool editandoEnBarraFormulas = false;
        public Form1()
        {
            InitializeComponentCustom();
            celdas = new Dictionary<string, string>();

            gestorArchivos = new GestorArchivos(dgvHoja, celdas, this);
            gestorSeleccion = new GestorSeleccion(dgvHoja);
            barraEstado = new BarraEstado(dgvHoja);

            this.Controls.Add(barraEstado.StatusStrip);

        }
        public void SimularAbrirArchivo()
        {
            this.BeginInvoke(new Action(() =>
            {
                System.Threading.Thread.Sleep(100);
                gestorArchivos.AbrirArchivo();
            }));
        }
        public string ArchivoActual
        {
            get => archivoActual;
            set => archivoActual = value;
        }
        public Dictionary<string, string> Celdas => celdas;
        public DataGridView DgvHoja => dgvHoja;
        public System.Windows.Forms.TextBox TxtFormula => txtFormula;
        public Label LblCelda => lblCelda;
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
            ConfigurarEventos();
            dgvHoja.BringToFront();

            barraEstado = new BarraEstado(dgvHoja);
            this.Controls.Add(barraEstado.StatusStrip);
            barraEstado.StatusStrip.Dock = DockStyle.Bottom;
            gestorSeleccion = new GestorSeleccion(dgvHoja);

        }
        private void CrearMenu()
        {
            menuStrip = new MenuStrip();
            menuStrip.BackColor = Color.FromArgb(250, 250, 250);
            menuStrip.Font = new Font("Segoe UI", 9F);

            var menuArchivo = new ToolStripMenuItem("Archivo");
            menuArchivo.DropDownItems.Add("Nuevo", null, (s, e) => gestorArchivos.NuevoArchivo());
            menuArchivo.DropDownItems.Add("Abrir", null, (s, e) => gestorArchivos.AbrirArchivo());
            menuArchivo.DropDownItems.Add("Guardar", null, (s, e) => gestorArchivos.GuardarArchivo());
            menuArchivo.DropDownItems.Add("Guardar Como", null, (s, e) => gestorArchivos.GuardarComo());
            menuArchivo.DropDownItems.Add(new ToolStripSeparator());
            menuArchivo.DropDownItems.Add("Salir", null, (s, e) => this.Close());

            var menuEdicion = new ToolStripMenuItem("Edición");
            menuEdicion.DropDownItems.Add("Copiar", null, (s, e) => gestorArchivos.Copiar());
            menuEdicion.DropDownItems.Add("Pegar", null, (s, e) => gestorArchivos.Pegar());
            menuEdicion.DropDownItems.Add("Cortar", null, (s, e) => gestorArchivos.Cortar());

            var menuInsertar = new ToolStripMenuItem("Insertar");
            menuInsertar.DropDownItems.Add("Fila", null, (s, e) => gestorArchivos.InsertarFila());
            menuInsertar.DropDownItems.Add("Columna", null, (s, e) => gestorArchivos.InsertarColumna());

            var menuFormulas = new ToolStripMenuItem("Fórmulas");
            menuFormulas.DropDownItems.Add("Suma", null, (s, e) => Formulas.InsertarFormula("SUMA", dgvHoja));
            menuFormulas.DropDownItems.Add("Resta", null, (s, e) => Formulas.InsertarFormula("RESTA", dgvHoja));
            menuFormulas.DropDownItems.Add("Multiplicar", null, (s, e) => Formulas.InsertarFormula("MULTIPLICAR", dgvHoja));
            menuFormulas.DropDownItems.Add("Dividir", null, (s, e) => Formulas.InsertarFormula("DIVIDIR", dgvHoja));

            menuFormulas.DropDownItems.Add("Promedio", null, (s, e) => Formulas.InsertarFormula("PROMEDIO", dgvHoja));
            menuFormulas.DropDownItems.Add("Máximo", null, (s, e) => Formulas.InsertarFormula("MAX", dgvHoja));
            menuFormulas.DropDownItems.Add("Mínimo", null, (s, e) => Formulas.InsertarFormula("MIN", dgvHoja));
            menuFormulas.DropDownItems.Add("Contar", null, (s, e) => Formulas.InsertarFormula("COUNT", dgvHoja));


            menuStrip.Items.AddRange(new ToolStripItem[] { menuArchivo, menuEdicion, menuInsertar, menuFormulas });
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }
        private void CrearToolbar()
        {
            toolStrip = new ToolStrip();
            toolStrip.BackColor = Color.FromArgb(245, 245, 245);

            var btnNuevo = new ToolStripButton("Nuevo");
            btnNuevo.Click += (s, e) => gestorArchivos.NuevoArchivo();

            var btnAbrir = new ToolStripButton("Abrir");
            btnAbrir.Click += (s, e) => gestorArchivos.AbrirArchivo();

            var btnGuardar = new ToolStripButton("Guardar");
            btnGuardar.Click += (s, e) => gestorArchivos.GuardarArchivo();

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
            txtFormula.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            panel.Controls.AddRange(new Control[] { lblCelda, txtFormula });
            this.Controls.Add(panel);
        }
        private void CrearDataGridView()
        {
            this.SuspendLayout();

            dgvHoja = new DataGridView();
            dgvHoja.Visible = false;
            dgvHoja.SuspendLayout();


            dgvHoja.Dock = DockStyle.Fill;
            dgvHoja.VirtualMode = false;
            dgvHoja.AllowUserToAddRows = false;
            dgvHoja.AllowUserToDeleteRows = false;
            dgvHoja.AllowUserToResizeRows = false;
            dgvHoja.RowHeadersWidth = 50;

            dgvHoja.ColumnHeadersDefaultCellStyle.Font = new Font("Calibri", 11F, FontStyle.Bold);
            dgvHoja.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvHoja.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(79, 129, 189);
            dgvHoja.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvHoja.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(198, 217, 191);
            dgvHoja.RowHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvHoja.RowHeadersDefaultCellStyle.Font = new Font("Calibri", 11F, FontStyle.Bold);
            dgvHoja.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            ((ISupportInitialize)dgvHoja).BeginInit();

            try
            {
                var columnasArray = new DataGridViewColumn[26];
                for (int i = 0; i < 26; i++)
                {
                    columnasArray[i] = new DataGridViewTextBoxColumn
                    {
                        Name = ((char)('A' + i)).ToString(),
                        HeaderText = ((char)('A' + i)).ToString(),
                        Width = 80,
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        Resizable = DataGridViewTriState.True
                    };
                }

                dgvHoja.Columns.AddRange(columnasArray);

                dgvHoja.RowCount = 100;
                dgvHoja.RowHeadersVisible = true;

                for (int i = 0; i < dgvHoja.RowCount; i++)
                {
                    dgvHoja.Rows[i].HeaderCell.Value = (i + 1).ToString();
                    dgvHoja.Rows[i].Height = 20;
                }

                dgvHoja.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                dgvHoja.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
                dgvHoja.DefaultCellStyle.Font = new Font("Calibri", 10F);
                dgvHoja.GridColor = Color.FromArgb(208, 215, 229);

                this.Controls.Add(dgvHoja);
            }
            finally
            {

                ((ISupportInitialize)dgvHoja).EndInit();
                dgvHoja.ResumeLayout();
                this.ResumeLayout(false);
            }

            HabilitarDoubleBuffering();

            Application.DoEvents();
            dgvHoja.Visible = true;
            dgvHoja.Refresh();
        }
        private void HabilitarDoubleBuffering()
        {
            if (dgvHoja != null)
            {
                try
                {
                    typeof(DataGridView).InvokeMember("DoubleBuffered",
                        BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                        null, dgvHoja, new object[] { true });
                }
                catch
                {

                }
            }
        }
        private void ConfigurarEventos()
        {
            dgvHoja.CellClick += DgvHoja_CellClick;
            dgvHoja.CellEndEdit += DgvHoja_CellEndEdit;
            dgvHoja.CellEnter += DgvHoja_CellEnter;
            dgvHoja.CellBeginEdit += DgvHoja_CellBeginEdit;
            dgvHoja.SelectionChanged += DgvHoja_SelectionChanged;
            dgvHoja.KeyDown += DgvHoja_KeyDown;
            txtFormula.KeyDown += TxtFormula_KeyDown;
            txtFormula.Enter += TxtFormula_Enter;
            txtFormula.Leave += TxtFormula_Leave;

            dgvHoja.MouseDown += (s, e) => gestorSeleccion.ManipularMouseDown(e);
            dgvHoja.MouseMove += (s, e) => gestorSeleccion.ManipularMouseMove(e);
            dgvHoja.MouseUp += (s, e) => gestorSeleccion.ManipularMouseUp(e);

            dgvHoja.Paint += DgvHoja_Paint;
        }
        private void DgvHoja_Paint(object sender, PaintEventArgs e)
        {
            gestorSeleccion.DibujarMarcoSeleccion(e.Graphics);
        }
        private void TxtFormula_Enter(object sender, EventArgs e)
        {
            editandoEnBarraFormulas = true;
            txtFormula.SelectAll();
        }
        private void TxtFormula_Leave(object sender, EventArgs e)
        {
            editandoEnBarraFormulas = false;
        }
        private void TxtFormula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (dgvHoja.CurrentCell != null)
                {
                    ProcesarContenidoCelda();
                    MoverACeldaSiguiente();
                    editandoEnBarraFormulas = false;
                    dgvHoja.Focus();
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                CancelarEdicion();
                e.Handled = true;
            }
        }
        private void ProcesarContenidoCelda()
        {
            var celda = dgvHoja.CurrentCell;
            string contenido = txtFormula.Text;
            string celdaRef = $"{(char)('A' + celda.ColumnIndex)}{celda.RowIndex + 1}";

            celdas[celdaRef] = contenido;

            if (contenido.StartsWith("="))
            {
                double resultado = Formulas.Evaluar(contenido, dgvHoja);
                celda.Tag = contenido; 

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
            else
            {
                celda.Value = contenido;
                celda.Tag = null;
                celda.Style.ForeColor = Color.Black;
            }

            dgvHoja.EndEdit();
        }
        private void MoverACeldaSiguiente()
        {
            var celda = dgvHoja.CurrentCell;
            int siguienteFila = celda.RowIndex + 1;

            if (siguienteFila < dgvHoja.Rows.Count)
            {
                dgvHoja.CurrentCell = dgvHoja.Rows[siguienteFila].Cells[celda.ColumnIndex];
                dgvHoja.ClearSelection();
                dgvHoja.CurrentCell.Selected = true;
            }
        }
        private void CancelarEdicion()
        {
            if (dgvHoja.CurrentCell != null)
            {
                txtFormula.Text = ObtenerContenidoOriginalCelda();
                editandoEnBarraFormulas = false;
                dgvHoja.Focus();
                dgvHoja.ClearSelection();
                dgvHoja.CurrentCell.Selected = true;
            }
        }
        private string ObtenerContenidoOriginalCelda()
        {
            var celda = dgvHoja.CurrentCell;
            string celdaRef = $"{(char)('A' + celda.ColumnIndex)}{celda.RowIndex + 1}";

            if (celdas.TryGetValue(celdaRef, out string valor))
                return valor;
            else if (celda.Tag is string formula)
                return formula;
            else
                return celda.Value?.ToString() ?? "";
        }
        private void DgvHoja_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                ActualizarBarraFormulas();
            }
        }
        private void DgvHoja_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!editandoEnBarraFormulas)
            {
                ActualizarBarraFormulas();
                barraEstado.Actualizar();
            }
        }
        private void DgvHoja_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            var celda = dgvHoja[e.ColumnIndex, e.RowIndex];
            if (celda.Tag is string formula && formula.StartsWith("="))
            {
                celda.Value = formula;
            }
        }
        private void DgvHoja_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            string celdaRef = $"{(char)('A' + e.ColumnIndex)}{e.RowIndex + 1}";
            string valor = dgvHoja[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";

            celdas[celdaRef] = valor;

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

            ActualizarBarraFormulas();
        }
        private void DgvHoja_SelectionChanged(object sender, EventArgs e)
        {
            string rangoSeleccionado = gestorSeleccion.ObtenerRangoSeleccionadoParaLabel();
            if (!string.IsNullOrEmpty(rangoSeleccionado))
            {
                lblCelda.Text = rangoSeleccionado;
            }
            barraEstado.Actualizar();
        }
        private void DgvHoja_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                if (dgvHoja.CurrentCell != null && !dgvHoja.IsCurrentCellInEditMode)
                {
                    LimpiarCeldasSeleccionadas();
                    e.Handled = true;
                }
            }
            else if (e.KeyCode == Keys.F2)
            {
                txtFormula.Focus();
                txtFormula.SelectAll();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                gestorSeleccion.LimpiarSeleccion();
                if (dgvHoja.CurrentCell != null)
                {
                    dgvHoja.CurrentCell.Selected = true;
                }
                e.Handled = true;
            }
        }
        private void LimpiarCeldasSeleccionadas()
        {
            if (gestorSeleccion.TieneSeleccionMultiple())
            {
                var celdasSeleccionadas = gestorSeleccion.ObtenerCeldasSeleccionadas();
                foreach (var celda in celdasSeleccionadas)
                {
                    LimpiarCeldaIndividual(celda);
                }
            }
            else if (dgvHoja.CurrentCell != null)
            {
                LimpiarCeldaIndividual(dgvHoja.CurrentCell);
            }

            ActualizarBarraFormulas();
            barraEstado.Actualizar();
        }
        public void ActualizarBarraFormulasConTexto(string texto)
        {
            txtFormula.Text = texto;
            txtFormula.Focus();
            txtFormula.SelectionStart = texto.Length;
        }
        private void LimpiarCeldaIndividual(DataGridViewCell celda)
        {
            string celdaRef = $"{(char)('A' + celda.ColumnIndex)}{celda.RowIndex + 1}";

            celda.Value = "";
            celda.Tag = null;
            celda.Style.ForeColor = Color.Black;

            if (celdas.ContainsKey(celdaRef))
                celdas.Remove(celdaRef);
        }
        private void ActualizarBarraFormulas()
        {
            if (dgvHoja.CurrentCell != null)
            {
                var celda = dgvHoja.CurrentCell;
                string celdaRef = $"{(char)('A' + celda.ColumnIndex)}{celda.RowIndex + 1}";

                lblCelda.Text = celdaRef;
                txtFormula.Text = ObtenerContenidoOriginalCelda();
            }
        }
        public void ActualizarTitulo(string nombreArchivo = null)
        {
            this.Text = string.IsNullOrEmpty(nombreArchivo) ?
                "Excel Básico - Nuevo archivo" :
                $"Excel Básico - {nombreArchivo}";
        }
        public void LimpiarHoja()
        {
            dgvHoja.Rows.Clear();
            celdas.Clear();
            txtFormula.Text = "";
            lblCelda.Text = "A1";

            dgvHoja.RowCount = 100;
            for (int i = 0; i < dgvHoja.RowCount; i++)
            {
                dgvHoja.Rows[i].HeaderCell.Value = (i + 1).ToString();
                dgvHoja.Rows[i].Height = 20;
            }
        }
    }
}