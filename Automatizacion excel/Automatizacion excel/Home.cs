using System;
using System.Windows.Forms;
using Automatizacion_excel.Formularios;
using Automatizacion_excel.RecortarExcel; // <--- ¡importante!

namespace Automatizacion_excel
{
    public partial class Home : Form
    {
        private Automatizacion_excel.Paso1.Paso1 paso1;
        private Automatizacion_excel.Paso2.Paso2 paso2;
        private Automatizacion_excel.Paso3.Paso3 paso3;
        private Automatizacion_excel.Paso4.Paso4 paso4;

        private Automatizacion_excel.QR.Paso1QR paso1QR;

        // Indica cuál flujo está activo
        private enum FlujoActivo { Ninguno, Fiserv, QR }
        private FlujoActivo flujoActual = FlujoActivo.Ninguno;

        // <--- BOTÓN NUEVO
        private Button btnRecortarMovimientos;

        public Home()
        {
            InitializeComponent();

            // --- Agregado: botón recortar movimientos al lado del IIBB
            btnRecortarMovimientos = new Button();
            btnRecortarMovimientos.Text = "Recortar Movimientos";
            btnRecortarMovimientos.Width = 180;
            btnRecortarMovimientos.Height = 40;
            btnRecortarMovimientos.Location = new System.Drawing.Point(
                btnVerIIBB.Right + 10, // a la par del botón IIBB, separadito
                btnVerIIBB.Top
            );
            btnRecortarMovimientos.Click += BtnRecortarMovimientos_Click;
            this.Controls.Add(btnRecortarMovimientos);
        }

        private void btnFiserv_Click(object sender, EventArgs e)
        {
            flujoActual = FlujoActivo.Fiserv;
            paso1 = new Automatizacion_excel.Paso1.Paso1(panelBotones, progressBar1, lblRutaArchivo, this);
            paso1.Paso1Completado += IniciarPaso2;
            panelBotones.Controls.Clear();
        }

        private void IniciarPaso2()
        {
            string ruta = paso1.ObtenerRutaExcel();
            paso2 = new Automatizacion_excel.Paso2.Paso2(panelBotones, progressBar1, lblRutaArchivo, this, ruta);
            paso2.Paso2Completado += IniciarPaso3;
        }

        private void IniciarPaso3(string rutaSas)
        {
            paso3 = new Automatizacion_excel.Paso3.Paso3(panelBotones, progressBar1, lblRutaArchivo, this, rutaSas);
            paso3.Paso3Completado += IniciarPaso4;
        }

        private void IniciarPaso4()
        {
            paso4 = new Automatizacion_excel.Paso4.Paso4(panelBotones, progressBar1, lblRutaArchivo, this);
            paso4.Ejecutar();
        }

        private void btnQR_Click(object sender, EventArgs e)
        {
            flujoActual = FlujoActivo.QR;
            paso1QR = new Automatizacion_excel.QR.Paso1QR(panelBotones, progressBar1, lblRutaArchivo, this);
            panelBotones.Controls.Clear();
        }

        private void btnSeleccionarArchivo_Click(object sender, EventArgs e)
        {
            if (flujoActual == FlujoActivo.Fiserv)
            {
                paso1?.SeleccionarArchivo();
            }
            else if (flujoActual == FlujoActivo.QR)
            {
                paso1QR?.SeleccionarArchivo();
            }
            else
            {
                MessageBox.Show("Primero elegí Fiserv o QR antes de seleccionar un archivo.");
            }
        }

        private void btnVerTasas_Click(object sender, EventArgs e)
        {
            var formTasas = new FormTasas();
            formTasas.ShowDialog();
        }

        private void btnVerIIBB_Click(object sender, EventArgs e)
        {
            var formIIBB = new FormIIBB();
            formIIBB.ShowDialog();
        }

        // --- AGREGADO: Lógica del nuevo botón
        private void BtnRecortarMovimientos_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";
            ofd.Title = "Seleccioná el Excel a recortar";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel|*.xlsx";
                sfd.Title = "Guardar recorte como...";
                sfd.FileName = "Recorte_movimientos.xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Ahora el método devuelve si copió movimientos o no
                        bool huboRecorte = Recortar_Excel.ProcesarArchivo(ofd.FileName, sfd.FileName);

                        if (huboRecorte)
                        {
                            MessageBox.Show("Archivo exportado correctamente a:\n" + sfd.FileName, "¡Éxito!");
                        }
                        else
                        {
                            // Intentar con el método alternativo (si lo tenés)
                            bool huboRecorte2 = Recortar_Excel2.ProcesarArchivo(ofd.FileName, sfd.FileName);

                            if (huboRecorte2)
                                MessageBox.Show("Archivo exportado con el método alternativo a:\n" + sfd.FileName, "¡Éxito (alternativo)!");
                            else
                                MessageBox.Show("No se encontraron movimientos para exportar con ninguno de los métodos.", "Aviso");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al recortar Excel:\n" + ex.Message, "Error");
                    }
                }
            }
        }




    }
}
