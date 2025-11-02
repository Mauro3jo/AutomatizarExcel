using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Automatizacion_excel.Paso2
{
    public class Paso2
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Label lblRutaSegundoArchivo;
        private Label lblEstadoProceso;
        private Form formularioPrincipal;

        private string rutaExcelPaso1;
        private string rutaExcelPaso2;

        private Button btnReubicarPorFecha;
        private Button btnPaso3;

        public event Action<string> Paso2Completado;

        public Paso2(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form, string rutaExcelAnterior)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            formularioPrincipal = form;
            rutaExcelPaso1 = rutaExcelAnterior;

            InicializarPaso2();
        }

        private void InicializarPaso2()
        {
            panelBotones.Controls.Clear();

            Label lblInfoOriginal = new Label
            {
                Text = $"📌 Archivo original cargado:\n{rutaExcelPaso1}",
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                Location = new Point(10, 10)
            };
            panelBotones.Controls.Add(lblInfoOriginal);

            Button btnCargarNuevo = new Button
            {
                Text = "📂 Cargar segundo Excel (SAS)",
                Width = 250,
                Height = 40,
                Location = new Point(10, 60)
            };
            btnCargarNuevo.Click += BtnCargarNuevo_Click;
            panelBotones.Controls.Add(btnCargarNuevo);

            lblRutaSegundoArchivo = new Label
            {
                Text = "",
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                Location = new Point(10, 110)
            };
            panelBotones.Controls.Add(lblRutaSegundoArchivo);

            btnReubicarPorFecha = new Button
            {
                Text = "⚙️ Procesar PENDIENTES-EXEP ANTICIPO",
                Width = 350,
                Height = 40,
                Location = new Point(10, 160),
                Enabled = false
            };
            btnReubicarPorFecha.Click += BtnReubicarPorFecha_Click;
            panelBotones.Controls.Add(btnReubicarPorFecha);

            lblEstadoProceso = new Label
            {
                Text = "⏳ Esperando acción...",
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                Location = new Point(10, 230)
            };
            panelBotones.Controls.Add(lblEstadoProceso);

            btnPaso3 = new Button
            {
                Text = "➡️ Seguir con el Paso 3 (manual)",
                Width = 250,
                Height = 40,
                Location = new Point(10, 270),
                Visible = true // se sigue mostrando
            };
            btnPaso3.Click += BtnPaso3_Click;
            panelBotones.Controls.Add(btnPaso3);

            progressBar.Visible = false;
            progressBar.Value = 0;
        }

        private void BtnCargarNuevo_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaExcelPaso2 = ofd.FileName;
                lblRutaSegundoArchivo.Text = $"📁 Segundo archivo cargado:\n{rutaExcelPaso2}";
                btnReubicarPorFecha.Enabled = true;
            }
        }

        private async void BtnReubicarPorFecha_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rutaExcelPaso2))
            {
                MessageBox.Show("Primero cargá el segundo archivo Excel (SAS).", "Falta archivo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var servicio = new ProcesarExcepcionAnticipoService();
                ActualizarEstado("🚀 Iniciando proceso completo...", 10);

                await System.Threading.Tasks.Task.Run(() =>
                {
                    servicio.EjecutarProceso(rutaExcelPaso2, ActualizarEstado);
                });

                MessageBox.Show("✔ Proceso completado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnPaso3.Visible = true;
            }
            catch (Exception ex)
            {
                ActualizarEstado("❌ Error inesperado: " + ex.Message, 0);
                MessageBox.Show("❌ Error al procesar operaciones:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPaso3_Click(object sender, EventArgs e)
        {
            var confirmar = MessageBox.Show(
                "Controlá el archivo SAS antes de continuar.\n¿Deseás seguir con el Paso 3 manualmente?",
                "Confirmación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmar == DialogResult.Yes && !string.IsNullOrEmpty(rutaExcelPaso2))
            {
                Paso2Completado?.Invoke(rutaExcelPaso2);
            }
        }

        private void ActualizarEstado(string mensaje, int progreso = -1)
        {
            try
            {
                // función local para tocar la UI
                void UpdateUI()
                {
                    if (lblEstadoProceso != null && !lblEstadoProceso.IsDisposed)
                        lblEstadoProceso.Text = mensaje ?? string.Empty;

                    if (progressBar != null && !progressBar.IsDisposed && progreso >= 0)
                    {
                        progressBar.Visible = true;
                        int val = Math.Min(Math.Max(progreso, 0), 100);
                        // evitar InvalidOperationException si Value == 100 y luego bajamos
                        if (val == 100)
                        {
                            progressBar.Value = 100;
                        }
                        else
                        {
                            if (progressBar.Value == 100) progressBar.Value = 0;
                            progressBar.Value = val;
                        }
                    }
                }

                // invocar en el hilo de UI si hace falta
                if (formularioPrincipal != null && formularioPrincipal.IsHandleCreated)
                {
                    if (formularioPrincipal.InvokeRequired)
                        formularioPrincipal.BeginInvoke((Action)(() => UpdateUI()));
                    else
                        UpdateUI();
                }
                else
                {
                    // fallback: si no hay handle aún, intentamos actualizar directo
                    UpdateUI();
                }
            }
            catch (ObjectDisposedException)
            {
                // La ventana ya se cerró; ignorar actualizaciones
            }
        }

    }
}
