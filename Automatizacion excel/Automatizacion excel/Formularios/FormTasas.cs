using System;
using System.Data;
using System.Windows.Forms;
using Automatizacion.Core.Formularios.Interfaz;
using Automatizacion.Core.Formularios.Servicios;
using Automatizacion.Core.Formularios.Validaciones;

namespace Automatizacion_excel.Formularios
{
    public partial class FormTasas : Form
    {
        private readonly ITasaService _tasaService;
        private DataTable tablaOriginal;
        private DataTable tablaEditada;
        private bool _hayCambios = false;

        public FormTasas()
        {
            InitializeComponent();
            _tasaService = new TasaService();
            CargarDatos();

            // Detectar cambios en la grilla
            dgvTasas.CellValueChanged += (s, e) => { _hayCambios = true; };
            dgvTasas.UserDeletedRow += (s, e) => { _hayCambios = true; };
            dgvTasas.UserAddedRow += (s, e) => { _hayCambios = true; };
            dgvTasas.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dgvTasas.IsCurrentCellDirty)
                    dgvTasas.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
        }

        private void CargarDatos()
        {
            try
            {
                tablaOriginal = _tasaService.ObtenerTasas();
                tablaEditada = tablaOriginal.Copy();
                dgvTasas.DataSource = tablaEditada;

                // El campo Id no debe ser editable, se ve gris y nunca editable
                if (dgvTasas.Columns.Contains("Id"))
                {
                    dgvTasas.Columns["Id"].ReadOnly = true;
                    dgvTasas.Columns["Id"].DefaultCellStyle.BackColor = System.Drawing.Color.LightGray;
                    dgvTasas.Columns["Id"].DefaultCellStyle.ForeColor = System.Drawing.Color.DarkGray;
                    dgvTasas.Columns["Id"].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.LightGray;
                    dgvTasas.Columns["Id"].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.DarkGray;
                }
            }
            catch (Exception ex)
            {
                LogError("Error al cargar tasas", ex);
                MessageBox.Show("❌ Error al cargar tasas: " + ex.Message);
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
                var errores = TasaValidator.ValidarTabla(tablaEditada);
                if (errores.Count > 0)
                {
                    MessageBox.Show("Hay errores en la tabla:\n" + string.Join("\n", errores),
                        "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }
                try
                {
                    _tasaService.GuardarTasas(tablaEditada);
                    MessageBox.Show("✔ Cambios guardados correctamente.");
                    base.OnFormClosing(e);
                }
                catch (Exception ex)
                {
                    LogError("Error al guardar tasas", ex);
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
            string ruta = "log_FormTasas.txt";
            System.IO.File.AppendAllText(ruta,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message} - {ex}\n");
        }
    }
}
