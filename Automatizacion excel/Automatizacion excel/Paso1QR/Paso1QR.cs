using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Automatizacion_excel.Paso1QR; // Namespace de tus servicios

namespace Automatizacion_excel.QR
{
    public class Paso1QR
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Form formularioPrincipal;
        private string rutaExcel; // El archivo principal (donde está QR y SAS)

        // Archivos CRM y Crudo
        private string rutaCRM;
        private string rutaCrudo;

        // UI labels para mostrar rutas
        private Label lblCRM;
        private Label lblCrudo;

        // Botones de acción sobre crudo
        private Button btnCopiarAltas;
        private Button btnCopiarBajas;
        private Button btnCopiarSas;

        // Botón NroAutorizacion (referencia, por si necesitás algo a futuro)
        private Button btnNroAutorizacion;

        public event Action Paso1QRCompletado;

        public Paso1QR(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            this.formularioPrincipal = form;
        }

        public void SeleccionarArchivo()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaExcel = ofd.FileName;
                lblRutaArchivo.Text = "Archivo QR/SAS: " + Path.GetFileName(rutaExcel);
                GenerarBotonesQR();
            }
        }

        private void GenerarBotonesQR()
        {
            panelBotones.Controls.Clear();

            int anchoLabel = 600;
            int posBotonX = 600;

            // ------ PRIMERA FILA: Botones QR ------  1036402 580385
            var btnProcesarQR = new Button
            {
                Text = "Procesar y pegar datos QR",
                Width = 200,
                Height = 36,
                Location = new Point(10, 10)
            };
            btnProcesarQR.Click += (s, e) => ProcesarYPegarQR();
            panelBotones.Controls.Add(btnProcesarQR);

            var btnEjecutarMacroQR = new Button
            {
                Text = "Ejecutar macro QR",
                Width = 170,
                Height = 36,
                Location = new Point(220, 10)
            };
            btnEjecutarMacroQR.Click += (s, e) =>
            {
                try
                {
                    MacroQR.EjecutarMacro(rutaExcel, "Limpiar");
                    MacroQR.EjecutarMacro(rutaExcel, "QR");
                    MessageBox.Show("Macros Limpiar y QR ejecutadas correctamente.", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al ejecutar las macros: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            panelBotones.Controls.Add(btnEjecutarMacroQR);

            btnNroAutorizacion = new Button
            {
                Text = "Ejecutar macro NroAutorizacion",
                Width = 210,
                Height = 36,
                Location = new Point(400, 10)
            };
            btnNroAutorizacion.Click += (s, e) =>
            {
                try
                {
                    MacroQR.EjecutarMacro(rutaExcel, "NroAutorizacion");
                    MessageBox.Show("Macro NroAutorizacion ejecutada correctamente.", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al ejecutar la macro NroAutorizacion: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            panelBotones.Controls.Add(btnNroAutorizacion);

            // ------ SEGUNDA FILA: Cargar archivos ------
            lblCRM = new Label
            {
                Text = "📁 CRM no cargado",
                AutoSize = false,
                Size = new Size(anchoLabel, 18),
                Location = new Point(10, 60)
            };
            panelBotones.Controls.Add(lblCRM);

            var btnCargarCRM = new Button
            {
                Text = "📂 Cargar archivo CRM",
                Width = 170,
                Height = 28,
                Location = new Point(posBotonX, 56)
            };
            btnCargarCRM.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    rutaCRM = ofd.FileName;
                    lblCRM.Text = $"📁 CRM cargado: {Path.GetFileName(rutaCRM)}";
                    VerificarArchivosCargados();
                }
            };
            panelBotones.Controls.Add(btnCargarCRM);

            lblCrudo = new Label
            {
                Text = "📁 Crudo no cargado",
                AutoSize = false,
                Size = new Size(anchoLabel, 18),
                Location = new Point(10, 90)
            };
            panelBotones.Controls.Add(lblCrudo);

            var btnCargarCrudo = new Button
            {
                Text = "📂 Cargar archivo Crudo",
                Width = 170,
                Height = 28,
                Location = new Point(posBotonX, 86)
            };
            btnCargarCrudo.Click += (s, e) =>
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm" };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    rutaCrudo = ofd.FileName;
                    lblCrudo.Text = $"📁 Crudo cargado: {Path.GetFileName(rutaCrudo)}";
                    VerificarArchivosCargados();
                }
            };
            panelBotones.Controls.Add(btnCargarCrudo);

            // ------ TERCERA FILA: Botones de copiar ------
            btnCopiarAltas = new Button
            {
                Text = "📋 Copiar ALTAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(10, 130),
                Enabled = false
            };
            btnCopiarAltas.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(rutaCRM) && !string.IsNullOrEmpty(rutaCrudo))
                {
                    var servicio = new CopiarAltasCrudoService();
                    EjecutarConProgreso(() => servicio.CopiarAltas(rutaCRM, rutaCrudo, Reportar));
                    MessageBox.Show("✔ Altas copiadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            panelBotones.Controls.Add(btnCopiarAltas);

            btnCopiarBajas = new Button
            {
                Text = "📋 Copiar BAJAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(270, 130),
                Enabled = false
            };
            btnCopiarBajas.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(rutaCRM) && !string.IsNullOrEmpty(rutaCrudo))
                {
                    var servicio = new CopiarBajasCrudoService();
                    EjecutarConProgreso(() => servicio.CopiarBajas(rutaCRM, rutaCrudo, Reportar));
                    MessageBox.Show("✔ Bajas copiadas correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            panelBotones.Controls.Add(btnCopiarBajas);

            btnCopiarSas = new Button
            {
                Text = "📋 Copiar SAS al Crudo",
                Width = 250,
                Height = 40,
                Location = new Point(530, 130),
                Enabled = false
            };
            btnCopiarSas.Click += (s, e) =>
            {
                if (!string.IsNullOrEmpty(rutaExcel) && !string.IsNullOrEmpty(rutaCrudo))
                {
                    var servicio = new CopiarSASCrudoService();
                    EjecutarConProgreso(() => servicio.CopiarSAS(rutaExcel, rutaCrudo, Reportar));
                    MessageBox.Show("✔ SAS copiado correctamente al Crudo.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            panelBotones.Controls.Add(btnCopiarSas);

            // Progreso inicial
            progressBar.Visible = false;
            progressBar.Value = 0;
        }

        private void VerificarArchivosCargados()
        {
            bool habilitar = !string.IsNullOrEmpty(rutaCRM) && !string.IsNullOrEmpty(rutaCrudo);
            btnCopiarAltas.Enabled = habilitar;
            btnCopiarBajas.Enabled = habilitar;
            btnCopiarSas.Enabled = !string.IsNullOrEmpty(rutaExcel) && !string.IsNullOrEmpty(rutaCrudo);
        }

        // QR original (igual que siempre)
        private void ProcesarYPegarQR()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Seleccioná el archivo desde donde copiar (A-V)";
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            string rutaOrigen = ofd.FileName;

            progressBar.Visible = true;
            progressBar.Value = 10;
            Application.DoEvents();

            bool ok = CopiarPegarEnHojaQR(rutaExcel, rutaOrigen);

            progressBar.Value = 100;
            progressBar.Visible = false;

            if (ok)
            {
                MessageBox.Show("¡Datos pegados en la hoja QR!", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Paso1QRCompletado?.Invoke();
            }
            else
            {
                MessageBox.Show("Ocurrió un error al procesar los archivos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CopiarPegarEnHojaQR(string rutaDestino, string rutaOrigen)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            Excel.Workbook wbDestino = null;
            Excel.Workbook wbOrigen = null;
            try
            {
                wbDestino = excelApp.Workbooks.Open(rutaDestino);
                excelApp.Run("LimpiarQR");
                wbDestino.Save();
                wbDestino.Close(false);
                Marshal.ReleaseComObject(wbDestino);
                wbDestino = null;

                wbDestino = excelApp.Workbooks.Open(rutaDestino);
                wbOrigen = excelApp.Workbooks.Open(rutaOrigen);

                Excel.Worksheet hojaDestino = wbDestino.Sheets["QR"] as Excel.Worksheet;
                Excel.Worksheet hojaOrigen = wbOrigen.Sheets[1] as Excel.Worksheet;

                int ultimaFilaOrigen = hojaOrigen.Cells[hojaOrigen.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                if (ultimaFilaOrigen < 2)
                    return false;

                Excel.Range rangoOrigen = hojaOrigen.Range["A2", hojaOrigen.Cells[ultimaFilaOrigen, 22]];
                int filasACopiar = rangoOrigen.Rows.Count;
                hojaDestino.Range[hojaDestino.Cells[2, 8], hojaDestino.Cells[filasACopiar + 1, 8]].NumberFormat = "@";
                Excel.Range rangoDestino = hojaDestino.Range[hojaDestino.Cells[2, 1], hojaDestino.Cells[1 + filasACopiar, 22]];
                rangoDestino.Value = rangoOrigen.Value;

                wbDestino.Save();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (wbOrigen != null) { wbOrigen.Close(false); Marshal.ReleaseComObject(wbOrigen); }
                if (wbDestino != null) { wbDestino.Close(); Marshal.ReleaseComObject(wbDestino); }
                excelApp.Quit(); Marshal.ReleaseComObject(excelApp);
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

        public string ObtenerRutaExcel() => rutaExcel;
    }
}
