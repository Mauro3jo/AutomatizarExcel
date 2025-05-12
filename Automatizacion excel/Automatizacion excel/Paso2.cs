using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Automatizacion_excel
{
    public class Paso2
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Label lblRutaSegundoArchivo;
        private Form formularioPrincipal;

        private string rutaExcelPaso1;
        private string rutaExcelPaso2;

        private DateTimePicker pickerFechaSeleccionada;
        private Button btnProcesarDesdeExcepcion;
        private Panel panelOpcional;

        public Paso2(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form, string rutaExcelAnterior)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            this.formularioPrincipal = form;
            this.rutaExcelPaso1 = rutaExcelAnterior;

            InicializarPaso2();
        }

        private void InicializarPaso2()
        {
            panelBotones.Controls.Clear();

            // Mostrar archivo original cargado
            Label lblInfoOriginal = new Label();
            lblInfoOriginal.Text = $"📌 Archivo original cargado:\n{rutaExcelPaso1}";
            lblInfoOriginal.AutoSize = true;
            lblInfoOriginal.MaximumSize = new Size(700, 0);
            lblInfoOriginal.Location = new Point(10, 10);
            panelBotones.Controls.Add(lblInfoOriginal);

            // Botón para cargar el segundo archivo
            Button btnCargarNuevo = new Button();
            btnCargarNuevo.Text = "📂 Cargar segundo Excel";
            btnCargarNuevo.Width = 200;
            btnCargarNuevo.Height = 40;
            btnCargarNuevo.Location = new Point(10, 60);
            btnCargarNuevo.Click += BtnCargarNuevo_Click;
            panelBotones.Controls.Add(btnCargarNuevo);

            // Label para mostrar ruta del segundo archivo
            lblRutaSegundoArchivo = new Label();
            lblRutaSegundoArchivo.Text = "";
            lblRutaSegundoArchivo.AutoSize = true;
            lblRutaSegundoArchivo.MaximumSize = new Size(700, 0);
            lblRutaSegundoArchivo.Location = new Point(10, 110);
            panelBotones.Controls.Add(lblRutaSegundoArchivo);

            // Botón obligatorio: Reubicar operaciones por fecha
            Button btnReubicarPorFecha = new Button();
            btnReubicarPorFecha.Text = "🔁 Reubicar operaciones por fecha";
            btnReubicarPorFecha.Width = 300;
            btnReubicarPorFecha.Height = 40;
            btnReubicarPorFecha.Location = new Point(10, 150);
            btnReubicarPorFecha.Click += BtnReubicarPorFecha_Click;
            panelBotones.Controls.Add(btnReubicarPorFecha);

            // Botón para mostrar sección opcional
            Button btnMostrarOpcional = new Button();
            btnMostrarOpcional.Text = "⚙️ Opcional (solo para feriados, etc.)";
            btnMostrarOpcional.Width = 300;
            btnMostrarOpcional.Height = 40;
            btnMostrarOpcional.Location = new Point(10, 200);
            btnMostrarOpcional.Click += BtnMostrarOpcional_Click;
            panelBotones.Controls.Add(btnMostrarOpcional);

            // Panel con controles opcionales (inicialmente oculto)
            panelOpcional = new Panel();
            panelOpcional.Visible = false;
            panelOpcional.Location = new Point(10, 250);
            panelOpcional.Size = new Size(700, 300);
            panelBotones.Controls.Add(panelOpcional);

            // Dentro del panel opcional
            Label lblFecha = new Label();
            lblFecha.Text = "📅 Fecha (columna Z):";
            lblFecha.AutoSize = true;
            lblFecha.Location = new Point(0, 0);
            panelOpcional.Controls.Add(lblFecha);

            pickerFechaSeleccionada = new DateTimePicker();
            pickerFechaSeleccionada.Format = DateTimePickerFormat.Custom;
            pickerFechaSeleccionada.CustomFormat = "d/M/yyyy";
            pickerFechaSeleccionada.Width = 150;
            pickerFechaSeleccionada.Location = new Point(0, 30);
            panelOpcional.Controls.Add(pickerFechaSeleccionada);

            btnProcesarDesdeExcepcion = new Button();
            btnProcesarDesdeExcepcion.Text = "➕ Agregar operaciones desde excepción";
            btnProcesarDesdeExcepcion.Width = 300;
            btnProcesarDesdeExcepcion.Height = 40;
            btnProcesarDesdeExcepcion.Location = new Point(0, 80);
            btnProcesarDesdeExcepcion.Click += BtnProcesarDesdeExcepcion_Click;
            btnProcesarDesdeExcepcion.Enabled = false;
            panelOpcional.Controls.Add(btnProcesarDesdeExcepcion);
        }

        private void BtnCargarNuevo_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";

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
                var filas = servicio.GenerarFilasDesdeExcepcion(rutaExcelPaso1, fechaSeleccionada);

                if (filas.Count == 0)
                {
                    MessageBox.Show("No se encontraron filas con esa fecha en la hoja 'excepcion anticipo'.", "Sin resultados", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var confirmar = MessageBox.Show($"Se encontraron {filas.Count} filas. ¿Deseás agregarlas al SAS del archivo original y también al nuevo?",
                    "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (confirmar != DialogResult.Yes) return;

                servicio.AgregarFilasAlSAS(rutaExcelPaso1, filas, progressBar);
                servicio.AgregarFilasAlSAS(rutaExcelPaso2, filas, progressBar);

                MessageBox.Show("✔ Operaciones agregadas exitosamente en ambos archivos SAS.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
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
                servicio.ProcesarOperaciones(rutaExcelPaso2, fechaCorte, folderQuitadas.SelectedPath, folderAgregadas.SelectedPath);

                MessageBox.Show("✔ Operaciones reubicadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("❌ Error al procesar operaciones:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
