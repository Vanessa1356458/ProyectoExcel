using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public partial class FormularioInicio : Form
    {
        private PictureBox pictureBox;
        private Label lblTitulo;
        private Label lblSubtitulo;
        private Button btnNuevo;
        private Button btnAbrir;
        private Button btnSalir;
        public FormularioInicio()
        {
            Iniciar();
        }
        private void Iniciar()
            {
            this.Icon = new Icon(SystemIcons.Application, 1, 1);
            this.Size = new Size(600, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White; 
            this.FormBorderStyle = FormBorderStyle.FixedSingle; 
            this.MaximizeBox = true; 
            this.MinimizeBox = true;

            pictureBox = new PictureBox();
            pictureBox.Size = new Size(120, 120);
            pictureBox.Location = new Point((this.ClientSize.Width - pictureBox.Width) / 2, 60);
            pictureBox.BackColor = Color.Transparent;
            pictureBox.BorderStyle = BorderStyle.None;

            Bitmap excelLogo = new Bitmap(120, 120);
            using (Graphics g = Graphics.FromImage(excelLogo))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                using (SolidBrush brush = new SolidBrush(Color.FromArgb(16, 124, 16)))
                {
                    g.FillRectangle(brush, 10, 10, 100, 100);
                }

                using (SolidBrush whiteBrush = new SolidBrush(Color.White))
                {
                    g.FillRectangle(whiteBrush, 25, 25, 70, 70);
                }

                using (Pen pen = new Pen(Color.FromArgb(16, 124, 16), 1))
                {
                    for (int i = 25; i <= 95; i += 14)
                    {
                        g.DrawLine(pen, 25, i, 95, i);
                    }
                    for (int i = 25; i <= 95; i += 14)
                    {
                        g.DrawLine(pen, i, 25, i, 95);
                    }
                }

                using (Font font = new Font("Arial", 28, FontStyle.Bold))
                using (SolidBrush textBrush = new SolidBrush(Color.FromArgb(16, 124, 16)))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    g.DrawString("X", font, textBrush, new Rectangle(25, 25, 70, 70), sf);
                }
            }
            pictureBox.Image = excelLogo;

            // Título
            lblTitulo = new Label();
            lblTitulo.Text = "Microsoft Excel";
            lblTitulo.Font = new Font(FontFamily.GenericSansSerif, 24, FontStyle.Regular);
            lblTitulo.ForeColor = Color.FromArgb(68, 68, 68); 
            lblTitulo.AutoSize = true;
            lblTitulo.Location = new Point((this.ClientSize.Width - 250) / 2, 200);
            lblTitulo.BackColor = Color.Transparent;

            // Botón "Libro en blanco" - Estilo Excel moderno
            btnNuevo = new Button();
            btnNuevo.Text = "Libro en blanco";
            btnNuevo.Size = new Size(160, 45);
            btnNuevo.Location = new Point(120, 300);
            btnNuevo.BackColor = Color.FromArgb(0, 120, 215); // Azul moderno de Office
            btnNuevo.ForeColor = Color.White;
            btnNuevo.FlatStyle = FlatStyle.Flat;
            btnNuevo.FlatAppearance.BorderSize = 0;
            btnNuevo.Font = new Font("Segoe UI", 11);
            btnNuevo.Cursor = Cursors.Hand;
            btnNuevo.Click += BtnNuevo_Click;

            // Efecto hover para btnNuevo
            btnNuevo.MouseEnter += (s, e) => {
                btnNuevo.BackColor = Color.FromArgb(16, 110, 190);
            };
            btnNuevo.MouseLeave += (s, e) => {
                btnNuevo.BackColor = Color.FromArgb(0, 120, 215);
            };

            // Botón "Abrir archivo" - Estilo Excel moderno
            btnAbrir = new Button();
            btnAbrir.Text = "Abrir archivo";
            btnAbrir.Size = new Size(160, 45);
            btnAbrir.Location = new Point(300, 300);
            btnAbrir.BackColor = Color.FromArgb(242, 242, 242); // Gris claro
            btnAbrir.ForeColor = Color.FromArgb(68, 68, 68);
            btnAbrir.FlatStyle = FlatStyle.Flat;
            btnAbrir.FlatAppearance.BorderSize = 1;
            btnAbrir.FlatAppearance.BorderColor = Color.FromArgb(204, 204, 204);
            btnAbrir.Font = new Font("Segoe UI", 11);
            btnAbrir.Cursor = Cursors.Hand;
            btnAbrir.Click += BtnAbrir_Click;

            // Efecto hover para btnAbrir
            btnAbrir.MouseEnter += (s, e) => {
                btnAbrir.BackColor = Color.FromArgb(232, 232, 232);
            };
            btnAbrir.MouseLeave += (s, e) => {
                btnAbrir.BackColor = Color.FromArgb(242, 242, 242);
            };

            // Agregar controles al formulario
            this.Controls.Add(pictureBox);
            this.Controls.Add(lblTitulo);
            this.Controls.Add(lblSubtitulo);
            this.Controls.Add(btnNuevo);
            this.Controls.Add(btnAbrir);

            // Centrar elementos cuando se redimensiona la ventana
            this.Resize += (s, e) => {
                pictureBox.Location = new Point((this.ClientSize.Width - pictureBox.Width) / 2, 60);
                lblTitulo.Location = new Point((this.ClientSize.Width - lblTitulo.Width) / 2, 200);
                lblSubtitulo.Location = new Point((this.ClientSize.Width - lblSubtitulo.Width) / 2, 240);
                btnNuevo.Location = new Point((this.ClientSize.Width - 340) / 2, 300);
                btnAbrir.Location = new Point((this.ClientSize.Width - 340) / 2 + 180, 300);
            };
        }
        private void BtnNuevo_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 excelForm = new Form1();
            excelForm.FormClosed += (s, args) => this.Close();
            excelForm.Show();
        }
        private void BtnAbrir_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 excelForm = new Form1();
            excelForm.FormClosed += (s, args) => this.Close();
            excelForm.Show();
            excelForm.SimularAbrirArchivo();
        }
    }
}
