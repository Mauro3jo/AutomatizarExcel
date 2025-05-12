using System;
using System.Windows.Forms;

namespace Automatizacion_excel
{
    public partial class Home : Form
    {
        private Paso1 paso1;
        private Paso2 paso2;

        public Home()
        {
            InitializeComponent();

            paso1 = new Paso1(panelBotones, progressBar1, lblRutaArchivo, this);
            paso1.Paso1Completado += IniciarPaso2;
        }

        private void IniciarPaso2()
        {
            string ruta = paso1.ObtenerRutaExcel();
            paso2 = new Paso2(panelBotones, progressBar1, lblRutaArchivo, this, ruta);
        }

        private void btnSeleccionarArchivo_Click(object sender, EventArgs e)
        {
            paso1.SeleccionarArchivo();
        }
    }
}
