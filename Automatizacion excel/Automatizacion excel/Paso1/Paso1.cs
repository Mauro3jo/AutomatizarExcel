using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public class Paso1
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Form formularioPrincipal;
        private string rutaExcel;
        private Button btnPaso2;
        private Button btnBrutoFinal; // NUEVO: botón Bruto Final
        private double[] brutosGuardados = new double[9]; // NUEVO: para almacenar los brutos de cada hoja
        private string[] hojasDeseadas = { // Lo pasamos a propiedad de clase porque se va a usar varias veces
            "Visa debito", "Mastercard debito", "MAESTRO", "Visa", "Mastercard",
            "ARGENCARD", "AMEX FISERV", "CABAL", "AMEX_2"
        };

        public event Action Paso1Completado;

        public Paso1(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form)
        {
            this.panelBotones = panelBotones;
            this.progressBar = progressBar;
            this.lblRutaArchivo = lblRutaArchivo;
            formularioPrincipal = form;
        }

        public void SeleccionarArchivo()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rutaExcel = ofd.FileName;
                lblRutaArchivo.Text = "Archivo: " + Path.GetFileName(rutaExcel);
                GenerarBotones();
            }
        }

        // Utilidad para copiar al portapapeles
        private void CopiarAlPortapapeles(string texto)
        {
            try { Clipboard.SetText(texto); }
            catch { }
        }

        // Muestra un label flotante arriba de un control (label/botón), y lo elimina después de 1.2 segundos
        private void MostrarFlotante(Control control, string texto)
        {
            Form flotante = new Form();
            flotante.FormBorderStyle = FormBorderStyle.None;
            flotante.StartPosition = FormStartPosition.Manual;
            flotante.ShowInTaskbar = false;
            flotante.TopMost = true;
            flotante.BackColor = Color.LightYellow;
            flotante.AutoSize = true;
            flotante.AutoSizeMode = AutoSizeMode.GrowAndShrink;

            Label lbl = new Label();
            lbl.Text = texto;
            lbl.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            lbl.ForeColor = Color.Green;
            lbl.BackColor = Color.LightYellow;
            lbl.AutoSize = true;
            lbl.Padding = new Padding(8, 4, 8, 4);
            flotante.Controls.Add(lbl);

            // Posición: Justo arriba y centrado
            Point screenPos = control.PointToScreen(Point.Empty);
            int x = screenPos.X + (control.Width - lbl.PreferredWidth) / 2;
            int y = screenPos.Y - lbl.Height - 8;
            if (y < 0) y = screenPos.Y + control.Height + 8;

            flotante.Location = new Point(x, y);

            flotante.Show();
            var timer = new System.Windows.Forms.Timer();
            timer.Interval = 1000;
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                flotante.Close();
                flotante.Dispose();
                timer.Dispose();
            };
            timer.Start();
        }




        private void GenerarBotones()
        {
            panelBotones.Controls.Clear();

            for (int idx = 0; idx < hojasDeseadas.Length; idx++)
            {
                var hoja = hojasDeseadas[idx];
                var panel = new Panel();
                panel.Width = 270;
                panel.Height = 40;

                var btn = new Button();
                btn.Text = hoja;
                btn.Width = 100;
                btn.Height = 30;
                btn.Tag = hoja;
                btn.Location = new Point(0, 5);
                btn.Click += (s, e) => ProcesarHoja(hoja);

                var lblFilas = new Label();
                lblFilas.Text = "0 filas";
                lblFilas.Width = 60;
                lblFilas.Location = new Point(105, 10);
                lblFilas.Name = "lblFilas_" + hoja.Replace(" ", "_");
                lblFilas.TextAlign = ContentAlignment.MiddleLeft;

                var lblBruto = new Label();
                lblBruto.Text = "$0,00";
                lblBruto.Width = 100;
                lblBruto.Location = new Point(170, 10);
                lblBruto.Name = "lbl_" + hoja.Replace(" ", "_");
                lblBruto.TextAlign = ContentAlignment.MiddleRight;

                int idxCopia = idx; // Para closure
                lblBruto.DoubleClick += (s, e) =>
                {
                    CopiarAlPortapapeles(lblBruto.Text);
                    MostrarFlotante(lblBruto, "Bruto copiado");
                };

                panel.Controls.Add(btn);
                panel.Controls.Add(lblFilas);
                panel.Controls.Add(lblBruto);
                panelBotones.Controls.Add(panel);
            }

            // Botón para continuar al Paso 2
            btnPaso2 = new Button();
            btnPaso2.Text = "Seguir con el Paso 2";
            btnPaso2.Width = 200;
            btnPaso2.Height = 40;
            btnPaso2.Visible = false;
            btnPaso2.Click += BtnPaso2_Click;
            panelBotones.Controls.Add(btnPaso2);

            // Botón para verificar columna O
            var btnVerificarColumnaO = new Button();
            btnVerificarColumnaO.Text = "Verificar columna O (Visa y MC Crédito)";
            btnVerificarColumnaO.Width = 270;
            btnVerificarColumnaO.Height = 40;
            btnVerificarColumnaO.Click += (s, e) =>
            {
                VerificarColumnaO("Visa");
                VerificarColumnaO("Mastercard");
                VerificarColumnaO("ARGENCARD");
            };
            panelBotones.Controls.Add(btnVerificarColumnaO);

            // Botón Bruto Final
            btnBrutoFinal = new Button();
            btnBrutoFinal.Text = "Bruto Final: $0,00";
            btnBrutoFinal.Width = 320;
            btnBrutoFinal.Height = 40;
            btnBrutoFinal.Enabled = false;
            btnBrutoFinal.BackColor = Color.LightSteelBlue;
            btnBrutoFinal.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            btnBrutoFinal.Click += (s, e) =>
            {
                CopiarAlPortapapeles(btnBrutoFinal.Text.Replace("Bruto Final: ", ""));
                MostrarFlotante(btnBrutoFinal, "Bruto final copiado");
            };
            panelBotones.Controls.Add(btnBrutoFinal);

            // 🔵 Botón Controlar Tasas
            var btnControlarTasas = new Button();
            btnControlarTasas.Name = "btnControlarTasas";
            btnControlarTasas.Text = "Controlar tasas";
            btnControlarTasas.Width = 270;
            btnControlarTasas.Height = 40;
            btnControlarTasas.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            btnControlarTasas.Click += BtnControlarTasas_Click;
            panelBotones.Controls.Add(btnControlarTasas);

            // Labels de resultado (Visa, MC, ArgenCard)
            var lblResultadoVisa = new Label();
            lblResultadoVisa.Name = "lblResultadoVisa";
            lblResultadoVisa.Width = 600;
            lblResultadoVisa.Location = new Point(10, panelBotones.Controls.Count * 45);
            panelBotones.Controls.Add(lblResultadoVisa);

            var lblResultadoMC = new Label();
            lblResultadoMC.Name = "lblResultadoMC";
            lblResultadoMC.Width = 600;
            lblResultadoMC.Location = new Point(10, (panelBotones.Controls.Count) * 45);
            panelBotones.Controls.Add(lblResultadoMC);

            var lblResultadoArg = new Label();
            lblResultadoArg.Name = "lblResultadoArg";
            lblResultadoArg.Width = 600;
            lblResultadoArg.Location = new Point(10, (panelBotones.Controls.Count) * 45);
            panelBotones.Controls.Add(lblResultadoArg);
        }


        // Llama cada vez que se actualiza un bruto
        private void ActualizarLabelYPanel(string hoja, double total)
        {
            // Actualizar el label de bruto y guardar el valor para la suma final
            var label = formularioPrincipal.Controls.Find("lbl_" + hoja.Replace(" ", "_"), true).FirstOrDefault() as Label;
            int idx = Array.IndexOf(hojasDeseadas, hoja);
            if (label != null && idx != -1)
            {
                label.Text = $"${total:N2}";
                brutosGuardados[idx] = total; // GUARDAMOS PARA EL FINAL
                var panel = label.Parent;
                bool yaTieneTilde = panel.Controls.OfType<Label>().Any(c => c.Text == "✔");
                if (!yaTieneTilde)
                {
                    var tilde = new Label();
                    tilde.Text = "✔";
                    tilde.ForeColor = Color.Green;
                    tilde.Location = new Point(250, 10);
                    tilde.Width = 20;
                    panel.Controls.Add(tilde);
                }
            }

            // Refrescar el total del Bruto Final
            ActualizarBrutoFinal();

            VerificarPaso1Completo();
        }

        // Suma los brutos y actualiza el botón Bruto Final
        private void ActualizarBrutoFinal()
        {
            double total = 0;
            foreach (var valor in brutosGuardados)
                total += valor;

            if (btnBrutoFinal != null)
            {
                btnBrutoFinal.Text = $"Bruto Final: ${total:N2}";
                btnBrutoFinal.Enabled = total > 0;
            }
        }
        private void ActualizarLabelFilas(string hoja, int filas)
        {
            var label = formularioPrincipal.Controls.Find("lblFilas_" + hoja.Replace(" ", "_"), true).FirstOrDefault() as Label;
            if (label != null)
            {
                label.Text = $"{filas} filas";
            }
        }

        private void VerificarPaso1Completo()
        {
            bool todasListas = panelBotones.Controls
                .OfType<Panel>()
                .All(p => p.Controls.OfType<Label>().Any(l => l.Text == "✔"));

            if (todasListas && btnPaso2 != null)
            {
                btnPaso2.Visible = true;
            }
        }

        private void BtnPaso2_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "¿Confirmás que revisaste y procesaste correctamente todas las hojas del Excel?",
                "Confirmación",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                panelBotones.Controls.Clear();
                Paso1Completado?.Invoke();
            }
        }

        private void MostrarBarra(bool visible)
        {
            progressBar.Visible = visible;
            progressBar.Value = visible ? 0 : 0;
        }
        private void ProcesarHoja(string hoja)
        {
            double total = 0;

            switch (hoja)
            {
                case "AMEX_2":
                    MostrarBarra(true);
                    var dtFilas = Amex_2Processor.ObtenerFilasCandidatas(rutaExcel, hoja, progressBar);
                    MostrarBarra(false);

                    if (dtFilas.Rows.Count == 0)
                    {
                        MostrarBarra(true);
                        int filasContadas;
                        total = Amex_2Processor.Procesar(rutaExcel, hoja, new List<int>(), progressBar, out filasContadas);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasContadas);
                        MessageBox.Show("No se encontraron filas vacías en la columna A. Se sumaron los montos directamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                        break;
                    }

                    var form = new VistaPreviaFilasForm(dtFilas);
                    form.MensajeAyuda = "✔ Se eliminarán las filas donde la columna A esté vacía.";

                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        int filasContadas;
                        total = Amex_2Processor.Procesar(rutaExcel, hoja, form.FilasSeleccionadas, progressBar, out filasContadas);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasContadas);
                        MessageBox.Show("Procesamiento finalizado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "CABAL":
                    MostrarBarra(true);
                    int filasCabal;
                    total = CabalProcessor.Procesar(rutaExcel, hoja, progressBar, out filasCabal);
                    ActualizarLabelYPanel(hoja, total);
                    ActualizarLabelFilas(hoja, filasCabal);
                    MessageBox.Show("CABAL procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarBarra(false);
                    break;

                case "AMEX FISERV":
                    MostrarBarra(true);
                    int filasFiserv;
                    total = AmexFiservProcessor.Procesar(rutaExcel, hoja, progressBar, out filasFiserv);
                    ActualizarLabelYPanel(hoja, total);
                    ActualizarLabelFilas(hoja, filasFiserv);
                    MessageBox.Show("AMEX FISERV procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarBarra(false);
                    break;

                case "ARGENCARD":
                    MostrarBarra(true);
                    var dtArg = ArgencardProcessor.ObtenerFilasAfectadas(rutaExcel, hoja, progressBar);
                    MostrarBarra(false);

                    if (dtArg.Rows.Count == 0)
                    {
                        MostrarBarra(true);
                        int filasArg;
                        total = ArgencardProcessor.Procesar(rutaExcel, hoja, new List<int>(), progressBar, out filasArg);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasArg);
                        MessageBox.Show("No se encontraron datos para limpiar. Se sumaron los montos directamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                        break;
                    }

                    var formArg = new VistaPreviaFilasForm(dtArg);
                    formArg.MensajeAyuda = "✔ Las filas que no empiezan con 01/ serán eliminadas.";

                    if (formArg.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        int filasArg;
                        total = ArgencardProcessor.Procesar(rutaExcel, hoja, formArg.FilasSeleccionadas, progressBar, out filasArg);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasArg);
                        MessageBox.Show("ARGENCARD procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;
                case "Mastercard":
                    MostrarBarra(true);
                    var dtMC = MastercardCreditoProcessor.ObtenerFilasAfectadas(rutaExcel, hoja);
                    MostrarBarra(false);

                    if (dtMC.Rows.Count == 0)
                    {
                        MostrarBarra(true);
                        int filasMC;
                        total = MastercardCreditoProcessor.Procesar(rutaExcel, hoja, new List<int>(), progressBar, out filasMC);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasMC);
                        MessageBox.Show("No se encontraron datos para filtrar. Se sumaron los montos directamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                        break;
                    }

                    var formMC = new VistaPreviaFilasForm(dtMC);
                    formMC.MensajeAyuda = "✔ Cuotas 3 y 6 se renombran. 02/06 se elimina. Montos multiplicados.";

                    if (formMC.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        int filasMC;
                        total = MastercardCreditoProcessor.Procesar(rutaExcel, hoja, formMC.FilasSeleccionadas, progressBar, out filasMC);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasMC);
                        MessageBox.Show("Mastercard Crédito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "Visa":
                    MostrarBarra(true);
                    var dtVisa = VisaCreditoProcessor.ObtenerFilasAfectadas(rutaExcel, hoja, progressBar);
                    MostrarBarra(false);

                    if (dtVisa.Rows.Count == 0)
                    {
                        MostrarBarra(true);
                        int filasVisa;
                        total = VisaCreditoProcessor.Procesar(rutaExcel, hoja, new List<int>(), progressBar, out filasVisa);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasVisa);
                        MessageBox.Show("No se encontraron datos para filtrar. Se sumaron los montos directamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                        break;
                    }

                    var formVisa = new VistaPreviaFilasForm(dtVisa);
                    formVisa.MensajeAyuda = "✔ Se eliminarán cuotas que no sean 01/XX. Montos se multiplican.";

                    if (formVisa.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        int filasVisa;
                        total = VisaCreditoProcessor.Procesar(rutaExcel, hoja, formVisa.FilasSeleccionadas, progressBar, out filasVisa);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasVisa);
                        MessageBox.Show("Visa Crédito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "Visa debito":
                    MostrarBarra(true);
                    int filasVisaDebito;
                    total = VisaDebitoProcessor.Procesar(rutaExcel, hoja, progressBar, out filasVisaDebito);
                    ActualizarLabelYPanel(hoja, total);
                    ActualizarLabelFilas(hoja, filasVisaDebito);
                    MessageBox.Show("Visa Débito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarBarra(false);
                    break;

                case "Mastercard debito":
                    MostrarBarra(true);
                    int filasMastercard;
                    total = MastercardDebitoProcessor.Procesar(rutaExcel, hoja, progressBar, out filasMastercard);
                    ActualizarLabelYPanel(hoja, total);
                    ActualizarLabelFilas(hoja, filasMastercard);
                    MessageBox.Show("Mastercard Débito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    MostrarBarra(false);
                    break;

                case "MAESTRO":
                    MostrarBarra(true);
                    int filasMaestro;
                    var dtMaestro = MaestroProcessor.ObtenerFilasAfectadas(rutaExcel, hoja, progressBar);
                    MostrarBarra(false);

                    if (dtMaestro.Rows.Count == 0)
                    {
                        MostrarBarra(true);
                        total = MaestroProcessor.Procesar(rutaExcel, hoja, new List<int>(), progressBar, out filasMaestro);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasMaestro);
                        MessageBox.Show("No se encontraron datos para limpiar. Se sumaron las ventas directamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                        break;
                    }

                    var formMaestro = new VistaPreviaFilasForm(dtMaestro);
                    formMaestro.MensajeAyuda = "✔ Se corregirán columnas desordenadas antes de sumar.";

                    if (formMaestro.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = MaestroProcessor.Procesar(rutaExcel, hoja, formMaestro.FilasSeleccionadas, progressBar, out filasMaestro);
                        ActualizarLabelYPanel(hoja, total);
                        ActualizarLabelFilas(hoja, filasMaestro);
                        MessageBox.Show("MAESTRO procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;
            }
        }

        private void VerificarColumnaO(string hoja)
        {
            string nombreVisible = hoja == "Visa" ? "Visa Crédito"
                                   : hoja == "Mastercard" ? "Mastercard Crédito"
                                   : hoja == "ARGENCARD" ? "ArgenCard"
                                   : hoja;

            string lblName = hoja == "Visa" ? "lblResultadoVisa"
                            : hoja == "Mastercard" ? "lblResultadoMC"
                            : hoja == "ARGENCARD" ? "lblResultadoArg"
                            : null;

            Label lblDestino = null;
            if (!string.IsNullOrEmpty(lblName))
            {
                lblDestino = formularioPrincipal.Controls.Find(lblName, true).FirstOrDefault() as Label;
            }

            try
            {
                // Mostrar progreso de inicio
                progressBar.Visible = true;
                progressBar.Value = 5;

                /*
                // Procesar y actualizar anticipos (ahora pasamos formularioPrincipal para los avisos)
                VerificarAnticipo.ProcesarYActualizarAnticipos(rutaExcel, hoja, formularioPrincipal);

                // Avanzar progreso
                progressBar.Value = 65;
                if (lblDestino != null)
                    lblDestino.Text = $"{nombreVisible}: Verificando filas sin anticipo...";
                */

                // Solo verificar si hay filas sin anticipo (columna O vacía)
                var filasVacias = VerificarAnticipo.FilasSinAnticipo(rutaExcel, hoja);

                progressBar.Value = 100;

                // Mostrar resultado en label
                if (lblDestino != null)
                {
                    if (filasVacias != null && filasVacias.Count > 0)
                        lblDestino.Text = $"{nombreVisible} sin anticipo: fila(s) {string.Join(", ", filasVacias)}";
                    else
                        lblDestino.Text = $"{nombreVisible} sin anticipo: ✔ OK";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al verificar anticipos en la hoja '{hoja}':\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (lblDestino != null)
                    lblDestino.Text = $"{nombreVisible}: Error al verificar.";
            }
            finally
            {
                progressBar.Value = 0;
                progressBar.Visible = false;
            }
        }


        private void BtnControlarTasas_Click(object sender, EventArgs e)
        {
            var tarjetas = new[] { "Visa", "Mastercard", "ARGENCARD" };
            var labelNames = new Dictionary<string, string>
    {
        { "Visa", "lblResultadoVisa" },
        { "Mastercard", "lblResultadoMC" },
        { "ARGENCARD", "lblResultadoArg" }
    };

            progressBar.Visible = true;
            progressBar.Value = 0;

            for (int idx = 0; idx < tarjetas.Length; idx++)
            {
                string tarjeta = tarjetas[idx];
                string nombreVisible = tarjeta == "Visa" ? "Visa Crédito"
                                         : tarjeta == "Mastercard" ? "Mastercard Crédito"
                                         : tarjeta == "ARGENCARD" ? "ArgenCard"
                                         : tarjeta;

                string lblName = labelNames[tarjeta];
                Label lblDestino = formularioPrincipal.Controls.Find(lblName, true).FirstOrDefault() as Label;
                if (lblDestino != null)
                    lblDestino.Text = $"{nombreVisible}: Analizando tasas...";

                progressBar.Value = 5 + idx * 30;

                // 1. Obtener tasas desde la base de datos!
                Dictionary<int, double> tasasPorCuota = ControlTasas.ObtenerTasasDesdeBD(tarjeta);

                // 2. Controlar las filas "Plan cuota"
                var filasConExceso = ControlTasas.VerificarExcesos(rutaExcel, tarjeta, tasasPorCuota);

                // 3. Mostrar resultado en el label
                if (lblDestino != null)
                {
                    if (filasConExceso.Count == 0)
                    {
                        lblDestino.Text = $"{nombreVisible}: ✔ OK";
                    }
                    else
                    {
                        lblDestino.Text = $"{nombreVisible}: tasas superadas en fila(s) {string.Join(", ", filasConExceso)}";
                    }
                }

                progressBar.Value = 35 + idx * 30;
            }

            progressBar.Value = 100;
            progressBar.Visible = false;
        }


        public string ObtenerRutaExcel()
        {
            return rutaExcel;
        }
    }
}
