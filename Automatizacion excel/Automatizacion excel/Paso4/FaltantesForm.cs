using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Automatizacion_excel.Paso4
{
    public partial class FaltantesForm : Form
    {
        private List<ControladorDiario.MiniFurRow> _faltantes;
        private string _rutaArchivoOriginal; // Si querés pasarla, sino no

        public FaltantesForm(List<ControladorDiario.MiniFurRow> faltantes)
        {
            InitializeComponent();
            _faltantes = faltantes;
        }

        private void FaltantesForm_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = _faltantes;
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel Files|*.xlsx";
            saveDialog.Title = "Guardar mini-FUR de faltantes";
            saveDialog.FileName = "mini_fur_faltantes.xlsx";

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                // Necesitás alguna ruta original para crear el ControladorDiario, o modificalo para no necesitar ruta
                var controlador = new ControladorDiario(_rutaArchivoOriginal ?? "");
                controlador.ExportarFaltantesAExcel(saveDialog.FileName, _faltantes);
                MessageBox.Show("Mini-FUR faltante exportado correctamente.");
            }
        }
    }
}
