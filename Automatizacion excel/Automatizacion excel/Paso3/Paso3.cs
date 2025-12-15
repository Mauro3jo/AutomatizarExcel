using System;
using System.Drawing;
using System.Windows.Forms;

namespace Automatizacion_excel.Paso3
{
    public class Paso3
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Form formularioPrincipal;
        private string rutaExcelPaso2;

        private string rutaExcelCRM;
        private string rutaExcelCrudo;

        private Label lblCRM;
        private Label lblCrudo;
        private Label lblSas;
        private Button btnCopiarAltas;
        private Button btnCopiarBajas;
        private Button btnCopiarSas;
        private Button btnPaso4;

        // Flags para verificar completitud
        private bool altasEjecutadas = false;
        private bool bajasEjecutadas = false;
        private bool sasEjecutado = false;

        // Evento que notifica a Home que puede iniciar Paso 4
        public event Action Paso3Completado;

        public Paso3(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form, string rutaExcelSAS)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            this.formularioPrincipal = form;
            this.rutaExcelPaso2 = rutaExcelSAS;

            InicializarPaso3();
        }

        private void InicializarPaso3()
        {
            panelBotones.Controls.Clear();

            int anchoLabel = 700;
            int posBotonX = 720;

            lblSas = new Label
            {
                Text = $"📄 Archivo SAS recibido:\n{rutaExcelPaso2}",
                AutoSize = false,
                Size = new Size(anchoLabel, 40),
                Location = new Point(10, 10)
            };
            panelBotones.Controls.Add(lblSas);

            lblCRM = new Label
            {
                Text = "📁 CRM no cargado",
                AutoSize = false,
                Size = new Size(anchoLabel, 20),
                Location = new Point(10, 60)
            };
            panelBotones.Controls.Add(lblCRM);

            var btnCargarCRM = new Button
            {
                Text = "📂 Cargar archivo CRM",
                Width = 160,
                Height = 30,
                Location = new Point(posBotonX, 55)
            };
            btnCargarCRM.Click += BtnCargarCRM_Click;
            panelBotones.Controls.Add(btnCargarCRM);

            lblCrudo = new Label
            {
                Text = "📁 Crudo no cargado",
                AutoSize = false,
                Size = new Size(anchoLabel, 20),
                Location = new Point(10, 100)
            };
            panelBotones.Controls.Add(lblCrudo);

            var btnCargarCrudo = new Button
            {
                Text = "📂 Cargar archivo Crudo",
                Width = 160,
                Height = 30,
                Location = new Point(posBotonX, 95)
            };
            btnCargarCrudo.Click += BtnCargarCrudo_Click;
            panelBotones.Controls.Add(btnCargarCrudo);

            btnCopiarAltas = new Button
            {
                Text = "📋 Copiar ALTAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(10, 140),
                Enabled = false
            };
            btnCopiarAltas.Click += BtnCopiarAltas_Click;
            panelBotones.Controls.Add(btnCopiarAltas);

            btnCopiarBajas = new Button
            {
                Text = "📋 Copiar BAJAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(270, 140),
                Enabled = false
            };
            btnCopiarBajas.Click += BtnCopiarBajas_Click;
            panelBotones.Controls.Add(btnCopiarBajas);

            btnCopiarSas = new Button
            {
                Text = "📋 Copiar SAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(530, 140),
                Enabled = !string.IsNullOrEmpty(rutaExcelPaso2)
            };
            btnCopiarSas.Click += BtnCopiarSas_Click;
            panelBotones.Controls.Add(btnCopiarSas);

            btnPaso4 = new Button
            {
                Text = "➡️ Seguir con el Paso 4",
                Width = 200,
                Height = 40,
                Location = new Point(10, 200),
                Visible = true
            };
            btnPaso4.Click += BtnPaso4_Click;
            panelBotones.Controls.Add(btnPaso4);
        }

        private void BtnCargarCRM_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaExcelCRM = ofd.FileName;
                lblCRM.Text = $"📁 CRM cargado: {rutaExcelCRM}";
                VerificarArchivosCargados();
            }
        }

        private void BtnCargarCrudo_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm" };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaExcelCrudo = ofd.FileName;
                lblCrudo.Text = $"📁 Crudo cargado: {rutaExcelCrudo}";
                VerificarArchivosCargados();
            }
        }

        private void VerificarArchivosCargados()
        {
            bool habilitar = !string.IsNullOrEmpty(rutaExcelCRM) && !string.IsNullOrEmpty(rutaExcelCrudo);
            btnCopiarAltas.Enabled = habilitar;
            btnCopiarBajas.Enabled = habilitar;
            btnCopiarSas.Enabled = habilitar && !string.IsNullOrEmpty(rutaExcelPaso2);
        }

        private void VerificarPasoCompletado()
        {
            if (altasEjecutadas && bajasEjecutadas && sasEjecutado)
            {
                btnPaso4.Visible = true;
            }
        }

        private void BtnCopiarAltas_Click(object sender, EventArgs e)
        {
            try
            {
                var servicio = new CopiarAltasDesdeCRMService();
                EjecutarConProgreso(() => servicio.CopiarAltas(rutaExcelCRM, rutaExcelCrudo, Reportar));
                MessageBox.Show("✔ Altas copiadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                altasEjecutadas = true;
                VerificarPasoCompletado();
            }
            catch (Exception ex)
            {
                MostrarError("ALTAS", ex);
            }
        }

        private void BtnCopiarBajas_Click(object sender, EventArgs e)
        {
            try
            {
                var servicio = new CopiarBajasDesdeCRMService();
                EjecutarConProgreso(() => servicio.CopiarBajas(rutaExcelCRM, rutaExcelCrudo, Reportar));
                MessageBox.Show("✔ Bajas copiadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bajasEjecutadas = true;
                VerificarPasoCompletado();
            }
            catch (Exception ex)
            {
                MostrarError("BAJAS", ex);
            }
        }

        private void BtnCopiarSas_Click(object sender, EventArgs e)
        {
            try
            {
                var servicio = new CopiarDesdeSASService();
                EjecutarConProgreso(() => servicio.CopiarSAS(rutaExcelPaso2, rutaExcelCrudo, Reportar));
                MessageBox.Show("✔ SAS copiado correctamente al Crudo.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sasEjecutado = true;
                VerificarPasoCompletado();
            }
            catch (Exception ex)
            {
                MostrarError("SAS", ex);
            }
        }

        private void BtnPaso4_Click(object sender, EventArgs e)
        {
            var confirmar = MessageBox.Show(
                "¿Deseás continuar con el Paso 4?",
                "Confirmación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmar == DialogResult.Yes)
            {
                Paso3Completado?.Invoke();
            }
        }

        private void EjecutarConProgreso(Action accion)
        {
            progressBar.Visible = true;
            progressBar.Value = 0;
            accion();
            progressBar.Visible = false;
            progressBar.Value = 0;
        }

        private void Reportar(string mensaje, int progreso)
        {
            lblRutaArchivo.Text = mensaje;
            progressBar.Value = Math.Min(Math.Max(progreso, 0), 100);
            Application.DoEvents();
        }

        private void MostrarError(string tipo, Exception ex)
        {
            MessageBox.Show($"❌ Error al copiar {tipo}:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
