using System;
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
        private Button btnDescargarResumen; // <--- Nuevo botón
        private Label lblResultado;

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

            // ------ Nuevo Botón: Exportar Resumen -------
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
            // --------------------------------------------

            lblResultado = new Label
            {
                Text = "",
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                Location = new Point(10, 130),
                ForeColor = Color.DarkBlue
            };
            panelBotones.Controls.Add(lblResultado);

            progressBar.Visible = false;
            progressBar.Value = 0;
        }

        private void BtnCargarDiario_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaDiario = ofd.FileName;
                lblDiario.Text = $"📁 Diario cargado:\n{rutaDiario}";
                btnControlarDiario.Enabled = true;
                btnValidarFUR.Enabled = true; // Habilita el botón FUR cuando se carga el diario
            }
        }

        private void BtnControlarDiario_Click(object sender, EventArgs e)
        {
            try
            {
                progressBar.Visible = true;
                progressBar.Value = 0;

                var controlador = new ControladorDiario(rutaDiario);

                // Paso 1: Validar Fecha
                string resultadoFecha = controlador.ControlarFechaUnica(out var filasInvalidas, ReportarProgreso);

                // Paso 2: Sumar Bruto
                string resultadoBruto = controlador.SumarColumnaBruto(ReportarProgreso);

                // Paso 3: Validar Arancel e IVA
                var (resultadoArancel, resultadoIva) = controlador.ValidarArancelEIVA(ReportarProgreso);

                // Paso 4: Controlar Costo Transaccional (AE)
                string resultadoCosto = controlador.ControlarCostoTransaccional(ReportarProgreso);

                // Paso 5: Validar IIBB
                var alicuotas = IIBBHelper.ObtenerAlicuotasDesdeBD();
                string resultadoIIBB = controlador.ValidarIIBB(ReportarProgreso, alicuotas);

                lblResultado.ForeColor = filasInvalidas.Count == 0 ? Color.Green : Color.Red;
                lblResultado.Text =
                    resultadoFecha + Environment.NewLine +
                    resultadoBruto + Environment.NewLine +
                    resultadoArancel + Environment.NewLine +
                    resultadoIva + Environment.NewLine +
                    resultadoCosto + Environment.NewLine +
                    resultadoIIBB;

                progressBar.Value = 100;
            }
            catch (Exception ex)
            {
                lblResultado.ForeColor = Color.Red;
                lblResultado.Text = "❌ Error: " + ex.Message;
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

                var controlador = new ControladorDiario(rutaDiario);
                string resultado = controlador.ValidarFUR(ReportarProgreso);

                lblResultado.ForeColor = resultado.StartsWith("✅") ? Color.Green : Color.Red;
                lblResultado.Text = resultado;

                progressBar.Value = 100;
            }
            catch (Exception ex)
            {
                lblResultado.ForeColor = Color.Red;
                lblResultado.Text = "❌ Error: " + ex.Message;
            }
            finally
            {
                progressBar.Visible = false;
                progressBar.Value = 0;
            }
        }

        // -------- NUEVO: Descargar Resumen --------
        private void BtnDescargarResumen_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rutaDiario) || !System.IO.File.Exists(rutaDiario))
            {
                MessageBox.Show("Primero cargá un archivo diario.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Archivos Excel|*.xlsx;*.xlsm",
                Title = "Guardar Resumen",
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
        // ------------------------------------------

        private void ReportarProgreso(string mensaje, int valor)
        {
            lblRutaArchivo.Text = mensaje;
            progressBar.Value = Math.Min(Math.Max(valor, 0), 100);
            Application.DoEvents();
        }
    }
}
