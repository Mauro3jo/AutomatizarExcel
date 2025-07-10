using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Automatizacion_excel
{
    public partial class VistaPreviaFilasForm : Form
    {
        public List<int> FilasSeleccionadas { get; private set; } = new List<int>();

        public string MensajeAyuda
        {
            set { lblAyuda.Text = value; }
        }
        public VistaPreviaFilasForm(DataTable filas)
        {
            InitializeComponent();
            CargarFilas(filas);
        }

        private void CargarFilas(DataTable dt)
        {
            var dtConSeleccion = dt.Copy();
            dtConSeleccion.Columns.Add("Aplicar", typeof(bool));

            foreach (DataRow row in dtConSeleccion.Rows)
                row["Aplicar"] = true;

            dgvFilas.DataSource = dtConSeleccion;
            dgvFilas.Columns["Aplicar"].DisplayIndex = 0;
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            FilasSeleccionadas.Clear();
            var dt = (DataTable)dgvFilas.DataSource;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dt.Rows[i]["Aplicar"]))
                {
                    int numeroFila = Convert.ToInt32(dt.Rows[i]["FilaExcel"]);
                    FilasSeleccionadas.Add(numeroFila);
                }
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
