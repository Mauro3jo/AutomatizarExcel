using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Automatizacion_excel
{
    public class Paso1
    {
        private Panel panelBotones;
        private ProgressBar progressBar;
        private Label lblRutaArchivo;
        private Form formularioPrincipal;
        private string rutaExcel;
        private Button btnPaso2;

        public event Action Paso1Completado;

        public Paso1(Panel panelBotones, ProgressBar progressBar, Label lblRutaArchivo, Form form)
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
                lblRutaArchivo.Text = "Archivo: " + Path.GetFileName(rutaExcel);
                GenerarBotones();
            }
        }

        private void GenerarBotones()
        {
            panelBotones.Controls.Clear();

            string[] hojasDeseadas = {
                "Visa debito", "Mastercard debito", "MAESTRO", "Visa", "Mastercard",
                "ARGENCARD", "AMEX FISERV", "CABAL", "AMEX_2"
            };

            foreach (var hoja in hojasDeseadas)
            {
                var panel = new Panel();
                panel.Width = 210;
                panel.Height = 40;

                var btn = new Button();
                btn.Text = hoja;
                btn.Width = 100;
                btn.Height = 30;
                btn.Tag = hoja;
                btn.Click += (s, e) => ProcesarHoja(hoja);

                var lbl = new Label();
                lbl.Text = "$0,00";
                lbl.Width = 70;
                lbl.Location = new System.Drawing.Point(105, 7);
                lbl.Name = "lbl_" + hoja.Replace(" ", "_");

                panel.Controls.Add(btn);
                panel.Controls.Add(lbl);
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
        }

        private void ProcesarHoja(string hoja)
        {
            double total = 0;

            switch (hoja)
            {
                case "AMEX_2":
                    var dtFilas = Amex_2Processor.ObtenerFilasCandidatas(rutaExcel, hoja);
                    var form = new VistaPreviaFilasForm(dtFilas);
                    form.MensajeAyuda = "✔ Se eliminarán las filas donde la columna A esté vacía.";

                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = Amex_2Processor.Procesar(rutaExcel, hoja, form.FilasSeleccionadas, progressBar);
                        ActualizarLabelYPanel(hoja, total);
                        MessageBox.Show("Procesamiento finalizado.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "CABAL":
                    total = CabalProcessor.Procesar(rutaExcel, hoja);
                    ActualizarLabelYPanel(hoja, total);
                    MessageBox.Show("CABAL procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "AMEX FISERV":
                    total = AmexFiservProcessor.Procesar(rutaExcel, hoja);
                    ActualizarLabelYPanel(hoja, total);
                    MessageBox.Show("AMEX FISERV procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "ARGENCARD":
                    var dtArg = ArgencardProcessor.ObtenerFilasAfectadas(rutaExcel, hoja);
                    var formArg = new VistaPreviaFilasForm(dtArg);
                    formArg.MensajeAyuda = "✔ Las filas que no empiezan con 01/ serán eliminadas.";

                    if (formArg.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = ArgencardProcessor.Procesar(rutaExcel, hoja, formArg.FilasSeleccionadas, progressBar);
                        ActualizarLabelYPanel(hoja, total);
                        MessageBox.Show("ARGENCARD procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "Mastercard":
                    var dtMC = MastercardCreditoProcessor.ObtenerFilasAfectadas(rutaExcel, hoja);
                    var formMC = new VistaPreviaFilasForm(dtMC);
                    formMC.MensajeAyuda = "✔ Cuotas 3 y 6 se renombran. 02/06 se elimina. Montos multiplicados.";

                    if (formMC.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = MastercardCreditoProcessor.Procesar(rutaExcel, hoja, formMC.FilasSeleccionadas, progressBar);
                        ActualizarLabelYPanel(hoja, total);
                        MessageBox.Show("Mastercard Crédito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "Visa":
                    var dtVisa = VisaCreditoProcessor.ObtenerFilasAfectadas(rutaExcel, hoja);
                    var formVisa = new VistaPreviaFilasForm(dtVisa);
                    formVisa.MensajeAyuda = "✔ Se eliminarán cuotas que no sean 01/XX. Montos se multiplican.";

                    if (formVisa.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = VisaCreditoProcessor.Procesar(rutaExcel, hoja, formVisa.FilasSeleccionadas, progressBar);
                        ActualizarLabelYPanel(hoja, total);
                        MessageBox.Show("Visa Crédito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;

                case "Visa debito":
                    total = VisaDebitoProcessor.Procesar(rutaExcel, hoja);
                    ActualizarLabelYPanel(hoja, total);
                    MessageBox.Show("Visa Débito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "Mastercard debito":
                    total = MastercardDebitoProcessor.Procesar(rutaExcel, hoja);
                    ActualizarLabelYPanel(hoja, total);
                    MessageBox.Show("Mastercard Débito procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                case "MAESTRO":
                    var dtMaestro = MaestroProcessor.ObtenerFilasAfectadas(rutaExcel, hoja);
                    var formMaestro = new VistaPreviaFilasForm(dtMaestro);
                    formMaestro.MensajeAyuda = "✔ Se corregirán columnas desordenadas antes de sumar.";

                    if (formMaestro.ShowDialog() == DialogResult.OK)
                    {
                        MostrarBarra(true);
                        total = MaestroProcessor.Procesar(rutaExcel, hoja, formMaestro.FilasSeleccionadas, progressBar);
                        ActualizarLabelYPanel(hoja, total);
                        MessageBox.Show("MAESTRO procesado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        MostrarBarra(false);
                    }
                    break;
            }
        }

        private void MostrarBarra(bool visible)
        {
            progressBar.Visible = visible;
            progressBar.Value = visible ? 0 : 0;
        }

        private void ActualizarLabelYPanel(string hoja, double total)
        {
            var label = formularioPrincipal.Controls.Find("lbl_" + hoja.Replace(" ", "_"), true).FirstOrDefault() as Label;
            if (label != null)
            {
                label.Text = $"${total:N2}";
                var panel = label.Parent;
                bool yaTieneTilde = panel.Controls.OfType<Label>().Any(c => c.Text == "✔");
                if (!yaTieneTilde)
                {
                    var tilde = new Label();
                    tilde.Text = "✔";
                    tilde.ForeColor = System.Drawing.Color.Green;
                    tilde.Location = new System.Drawing.Point(180, 7);
                    tilde.Width = 20;
                    panel.Controls.Add(tilde);
                }
            }

            VerificarPaso1Completo();
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

        public string ObtenerRutaExcel()
        {
            return rutaExcel;
        }
    }
}
