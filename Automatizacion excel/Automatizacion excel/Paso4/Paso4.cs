using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Automatizacion_excel.Paso4;

namespace Automatizacion_excel.Paso4
{
    public class Paso4
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Form formularioPrincipal;

        private string rutaDiario;
        private Label lblDiario;
        private Button btnCargarDiario;
        private Button btnControlarDiario;
        private Button btnValidarFUR;
        private Button btnDescargarResumen;

        // 🔥 CAMBIO CLAVE
        private RichTextBox rtbResultado;

        public Paso4(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            this.formularioPrincipal = form;
        }

        public void Ejecutar()
        {
            InicializarPaso4();
        }

        private void InicializarPaso4()
        {
            panelBotones.Controls.Clear();

            int anchoLabel = 700;
            int posBotonX = 720;

            lblDiario = new Label
            {
                Text = "📁 Diario no cargado",
                AutoSize = false,
                Size = new Size(anchoLabel, 40),
                Location = new Point(10, 20),
                Font = new Font("Segoe UI", 9)
            };
            panelBotones.Controls.Add(lblDiario);

            btnCargarDiario = new Button
            {
                Text = "📂 Cargar archivo DIARIO",
                Width = 160,
                Height = 30,
                Location = new Point(posBotonX, 25)
            };
            btnCargarDiario.Click += BtnCargarDiario_Click;
            panelBotones.Controls.Add(btnCargarDiario);

            btnControlarDiario = new Button
            {
                Text = "📋 Controlar Diario",
                Width = 200,
                Height = 40,
                Location = new Point(10, 80),
                Enabled = false
            };
            btnControlarDiario.Click += BtnControlarDiario_Click;
            panelBotones.Controls.Add(btnControlarDiario);

            btnValidarFUR = new Button
            {
                Text = "🧐 Validar FUR",
                Width = 160,
                Height = 30,
                Location = new Point(posBotonX + 180, 25),
                Enabled = false
            };
            btnValidarFUR.Click += BtnValidarFUR_Click;
            panelBotones.Controls.Add(btnValidarFUR);

            btnDescargarResumen = new Button
            {
                Text = "⬇️ Descargar Resumen",
                Width = 220,
                Height = 30,
                Location = new Point(10, 180),
                Enabled = true
            };
            btnDescargarResumen.Click += BtnDescargarResumen_Click;
            panelBotones.Controls.Add(btnDescargarResumen);

            // ✅ RESULTADOS CON COLOR POR LÍNEA
            rtbResultado = new RichTextBox
            {
                Location = new Point(10, 130),
                Size = new Size(700, 240),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9),
                BackColor = panelBotones.BackColor
            };
            panelBotones.Controls.Add(rtbResultado);

            progressBar.Visible = false;
            progressBar.Value = 0;
        }

        private void BtnCargarDiario_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaDiario = ofd.FileName;
                lblDiario.Text = $"📁 Diario cargado:\n{rutaDiario}";
                btnControlarDiario.Enabled = true;
                btnValidarFUR.Enabled = true;
            }
        }

        private void BtnControlarDiario_Click(object sender, EventArgs e)
        {
            progressBar.Visible = true;
            progressBar.Value = 0;

            rtbResultado.Clear();

            var controlador = new ControladorDiario(rutaDiario);

            string resultadoFecha;
            string resultadoBruto;
            string resultadoArancel;
            string resultadoIva;
            string resultadoCosto;
            string resultadoIIBB;

            List<int> filasInvalidas = new List<int>();

            try
            {
                resultadoFecha = controlador.ControlarFechaUnica(out filasInvalidas, ReportarProgreso);
                resultadoBruto = controlador.SumarColumnaBruto(ReportarProgreso);

                var (rArancel, rIva) = controlador.ValidarArancelEIVA(ReportarProgreso);
                resultadoArancel = rArancel;
                resultadoIva = rIva;

                resultadoCosto = controlador.ControlarCostoTransaccional(ReportarProgreso);

                var alicuotas = IIBBHelper.ObtenerAlicuotasDesdeBD();
                resultadoIIBB = controlador.ValidarIIBB(ReportarProgreso, alicuotas);

                AgregarLinea(resultadoFecha);
                AgregarLinea(resultadoBruto);
                AgregarLinea(resultadoArancel);
                AgregarLinea(resultadoIva);
                AgregarLinea(resultadoCosto);
                AgregarLinea(resultadoIIBB);

                progressBar.Value = 100;
            }
            catch (Exception ex)
            {
                AgregarLinea("❌ Error general: " + ex.Message);
            }
            finally
            {
                progressBar.Visible = false;
                progressBar.Value = 0;
            }
        }

        private void BtnValidarFUR_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar.Visible = true;
                progressBar.Value = 0;

                rtbResultado.Clear();

                var controlador = new ControladorDiario(rutaDiario);
                string resultado = controlador.ValidarFUR(ReportarProgreso);

                AgregarLinea(resultado);

                progressBar.Value = 100;
            }
            catch (Exception ex)
            {
                AgregarLinea("❌ Error: " + ex.Message);
            }
            finally
            {
                progressBar.Visible = false;
                progressBar.Value = 0;
            }
        }

        private void BtnDescargarResumen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rutaDiario))
            {
                MessageBox.Show("Primero cargá un archivo diario.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xlsm",
                FileName = "ResumenComisiones.xlsx"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var exportador = new ExportarExcel();
                    exportador.ExportarResumenComisiones(rutaDiario, sfd.FileName);
                    MessageBox.Show("¡Resumen exportado correctamente!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al exportar: " + ex.Message);
                }
            }
        }

        private void AgregarLinea(string texto)
        {
            rtbResultado.SelectionStart = rtbResultado.TextLength;
            rtbResultado.SelectionLength = 0;

            if (texto.StartsWith("✅"))
                rtbResultado.SelectionColor = Color.Green;
            else if (texto.StartsWith("❌"))
                rtbResultado.SelectionColor = Color.Red;
            else
                rtbResultado.SelectionColor = Color.Black;

            rtbResultado.AppendText(texto + Environment.NewLine);
            rtbResultado.SelectionColor = rtbResultado.ForeColor;
        }

        private void ReportarProgreso(string mensaje, int valor)
        {
            lblRutaArchivo.Text = mensaje;
            progressBar.Value = Math.Min(Math.Max(valor, 0), 100);
            Application.DoEvents();
        }
    }
}
