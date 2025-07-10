using System;
using System.Data;
using System.Windows.Forms;
using Automatizacion.Core.Formularios.Interfaz;
using Automatizacion.Core.Formularios.Servicios;
using Automatizacion.Core.Formularios.Validaciones;

namespace Automatizacion_excel.Formularios
{
    public partial class FormIIBB : Form
    {
        private readonly IProvinciaService _provinciaService;
        private DataTable tablaOriginal;
        private DataTable tablaEditada;
        private bool _hayCambios = false;

        public FormIIBB()
        {
            InitializeComponent();
            _provinciaService = new ProvinciaService();
            CargarDatos();

            // Detectar cambios en la grilla
            dgvIIBB.CellValueChanged += (s, e) => { _hayCambios = true; };
            dgvIIBB.UserDeletedRow += (s, e) => { _hayCambios = true; };
            dgvIIBB.UserAddedRow += (s, e) => { _hayCambios = true; };
            dgvIIBB.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvIIBB.IsCurrentCellDirty)
                    dgvIIBB.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
        }

        private void CargarDatos()
        {
            try
            {
                tablaOriginal = _provinciaService.ObtenerProvincias();
                tablaEditada = tablaOriginal.Copy();
                dgvIIBB.DataSource = tablaEditada;

                // El campo id no debe ser editable, se ve gris y nunca editable
                if (dgvIIBB.Columns.Contains("id"))
                {
                    dgvIIBB.Columns["id"].ReadOnly = true;
                    dgvIIBB.Columns["id"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                    dgvIIBB.Columns["id"].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkGray;
                    dgvIIBB.Columns["id"].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.LightGray;
                    dgvIIBB.Columns["id"].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.DarkGray;
                }
            }
            catch (Exception ex)
            {
                LogError("Error al cargar provincias", ex);
                MessageBox.Show("❌ Error al cargar provincias: " + ex.Message);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_hayCambios)
            {
                base.OnFormClosing(e);
                return;
            }

            var dr = MessageBox.Show(
                "Se detectaron cambios. ¿Desea guardar antes de salir?",
                "Guardar cambios",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                var errores = ProvinciaValidator.ValidarTabla(tablaEditada);
                if (errores.Count > 0)
                {
                    MessageBox.Show("Hay errores en la tabla:\n" + string.Join("\n", errores),
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                try
                {
                    _provinciaService.GuardarProvincias(tablaEditada);
                    MessageBox.Show("✔ Cambios guardados correctamente.");
                    base.OnFormClosing(e);
                }
                catch (Exception ex)
                {
                    LogError("Error al guardar provincias", ex);
                    MessageBox.Show("❌ Error al guardar: " + ex.Message);
                    e.Cancel = true;
                }
            }
            else
            {
                base.OnFormClosing(e);
            }
        }

        private void LogError(string message, Exception ex)
        {
            string ruta = "log_FormIIBB.txt";
            System.IO.File.AppendAllText(ruta,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message} - {ex}\n");
        }
    }
}
