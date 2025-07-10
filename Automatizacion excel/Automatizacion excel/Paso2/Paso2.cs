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

        private DateTimePicker pickerFechaSeleccionada;
        private Button btnProcesarDesdeExcepcion;
        private Button btnPaso3;
        private Panel panelOpcional;

        // 👉 EVENTO para que Home escuche
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
                Text = "📂 Cargar segundo Excel",
                Width = 200,
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

            Button btnReubicarPorFecha = new Button
            {
                Text = "🔁 Reubicar operaciones por fecha",
                Width = 300,
                Height = 40,
                Location = new Point(10, 150)
            };
            btnReubicarPorFecha.Click += BtnReubicarPorFecha_Click;
            panelBotones.Controls.Add(btnReubicarPorFecha);

            Button btnMostrarOpcional = new Button
            {
                Text = "⚙️ Opcional (solo para feriados, etc.)",
                Width = 300,
                Height = 40,
                Location = new Point(10, 200)
            };
            btnMostrarOpcional.Click += BtnMostrarOpcional_Click;
            panelBotones.Controls.Add(btnMostrarOpcional);

            panelOpcional = new Panel
            {
                Visible = false,
                Location = new Point(10, 250),
                Size = new Size(700, 300)
            };
            panelBotones.Controls.Add(panelOpcional);

            Label lblFecha = new Label
            {
                Text = "📅 Fecha (columna Z):",
                AutoSize = true,
                Location = new Point(0, 0)
            };
            panelOpcional.Controls.Add(lblFecha);

            pickerFechaSeleccionada = new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "d/M/yyyy",
                Width = 150,
                Location = new Point(0, 30)
            };
            panelOpcional.Controls.Add(pickerFechaSeleccionada);

            btnProcesarDesdeExcepcion = new Button
            {
                Text = "➕ Agregar operaciones desde excepción",
                Width = 300,
                Height = 40,
                Location = new Point(0, 80),
                Enabled = false
            };
            btnProcesarDesdeExcepcion.Click += BtnProcesarDesdeExcepcion_Click;
            panelOpcional.Controls.Add(btnProcesarDesdeExcepcion);

            lblEstadoProceso = new Label
            {
                Text = "⏳ Esperando acción...",
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                Location = new Point(10, 570)
            };
            panelBotones.Controls.Add(lblEstadoProceso);

            // Botón para seguir con Paso 3 (inicialmente oculto)
            btnPaso3 = new Button
            {
                Text = "➡️ Seguir con el Paso 3",
                Width = 200,
                Height = 40,
                Location = new Point(10, 620),
                Visible = false
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
                btnProcesarDesdeExcepcion.Enabled = true;
            }
        }

        private void BtnMostrarOpcional_Click(object sender, EventArgs e)
        {
            panelOpcional.Visible = !panelOpcional.Visible;
        }

        private void BtnProcesarDesdeExcepcion_Click(object sender, EventArgs e)
        {
            string fechaSeleccionada = pickerFechaSeleccionada.Value.ToString("d/M/yyyy");
            var servicio = new OperacionesDesdeExcepcionService();

            try
            {
                var filas = servicio.GenerarFilasDesdeExcepcion(rutaExcelPaso1, fechaSeleccionada, ActualizarEstado);

                if (filas.Count == 0)
                {
                    MessageBox.Show("No se encontraron filas con esa fecha en la hoja 'excepcion anticipo'.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var confirmar = MessageBox.Show($"Se encontraron {filas.Count} filas. ¿Deseás agregarlas al SAS del archivo original y también al nuevo?",
                    "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmar != DialogResult.Yes) return;

                servicio.AgregarFilasAlSAS(rutaExcelPaso1, filas, ActualizarEstado);
                servicio.AgregarFilasAlSAS(rutaExcelPaso2, filas, ActualizarEstado);

                MessageBox.Show("✔ Operaciones agregadas exitosamente en ambos archivos SAS.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ActualizarEstado("❌ Error inesperado: " + ex.Message, 0);
                MessageBox.Show("Ocurrió un error:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReubicarPorFecha_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rutaExcelPaso2))
            {
                MessageBox.Show("Primero cargá el segundo archivo Excel (SAS).", "Falta archivo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            FolderBrowserDialog folderQuitadas = new FolderBrowserDialog
            {
                Description = "Seleccioná la carpeta donde se guardarán las operaciones quitadas"
            };
            if (folderQuitadas.ShowDialog() != DialogResult.OK) return;

            FolderBrowserDialog folderAgregadas = new FolderBrowserDialog
            {
                Description = "Seleccioná la carpeta donde se guardarán las operaciones agregadas"
            };
            if (folderAgregadas.ShowDialog() != DialogResult.OK) return;

            DateTimePicker fechaPicker = new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "d/M/yyyy",
                Width = 200
            };
            var inputForm = new Form
            {
                Text = "Seleccioná la fecha máxima permitida",
                Width = 300,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen
            };
            fechaPicker.Location = new Point(30, 20);
            inputForm.Controls.Add(fechaPicker);
            Button btnOK = new Button { Text = "Aceptar", DialogResult = DialogResult.OK, Location = new Point(100, 60) };
            inputForm.Controls.Add(btnOK);
            inputForm.AcceptButton = btnOK;

            if (inputForm.ShowDialog() != DialogResult.OK) return;

            DateTime fechaCorte = fechaPicker.Value;

            try
            {
                var servicio = new OperacionesPorFechaService();
                servicio.ProcesarOperaciones(rutaExcelPaso2, fechaCorte, folderQuitadas.SelectedPath, folderAgregadas.SelectedPath, ActualizarEstado);

                MessageBox.Show("✔ Operaciones reubicadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

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
                "Antes de continuar, controlá el archivo SAS.\n¿Deseás seguir con el Paso 3?",
                "Confirmación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmar == DialogResult.Yes && !string.IsNullOrEmpty(rutaExcelPaso2))
            {
                // 👉 Emitir evento, que captará Home.cs
                Paso2Completado?.Invoke(rutaExcelPaso2);
            }
        }

        private void ActualizarEstado(string mensaje, int progreso = -1)
        {
            if (lblEstadoProceso != null)
                lblEstadoProceso.Text = mensaje;

            if (progressBar != null && progreso >= 0)
            {
                progressBar.Visible = true;
                progressBar.Value = Math.Min(Math.Max(progreso, 0), 100);
            }

            Application.DoEvents(); // Forzar UI update
        }
    }
}
