namespace Automatizacion_excel
{
    partial class VistaPreviaFilasForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView dgvFilas;
        private System.Windows.Forms.Button btnAceptar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Label lblAyuda;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.dgvFilas = new System.Windows.Forms.DataGridView();
            this.btnAceptar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.lblAyuda = new System.Windows.Forms.Label();

            ((System.ComponentModel.ISupportInitialize)(this.dgvFilas)).BeginInit();
            this.SuspendLayout();

            // lblAyuda
            this.lblAyuda.AutoSize = true;
            this.lblAyuda.Location = new System.Drawing.Point(12, 9);
            this.lblAyuda.Name = "lblAyuda";
            this.lblAyuda.Size = new System.Drawing.Size(0, 15);
            this.lblAyuda.TabIndex = 3;
            this.lblAyuda.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.lblAyuda.Text = "";

            // dgvFilas
            this.dgvFilas.AllowUserToAddRows = false;
            this.dgvFilas.AllowUserToDeleteRows = false;
            this.dgvFilas.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvFilas.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvFilas.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvFilas.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.dgvFilas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvFilas.Location = new System.Drawing.Point(12, 35);
            this.dgvFilas.Name = "dgvFilas";
            this.dgvFilas.Size = new System.Drawing.Size(760, 330);
            this.dgvFilas.TabIndex = 0;

            // btnAceptar
            this.btnAceptar.Location = new System.Drawing.Point(590, 380);
            this.btnAceptar.Size = new System.Drawing.Size(80, 30);
            this.btnAceptar.Text = "Aplicar";
            this.btnAceptar.UseVisualStyleBackColor = true;
            this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);

            // btnCancelar
            this.btnCancelar.Location = new System.Drawing.Point(680, 380);
            this.btnCancelar.Size = new System.Drawing.Size(80, 30);
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);

            // VistaPreviaFilasForm
            this.ClientSize = new System.Drawing.Size(784, 421);
            this.Controls.Add(this.lblAyuda);
            this.Controls.Add(this.dgvFilas);
            this.Controls.Add(this.btnAceptar);
            this.Controls.Add(this.btnCancelar);
            this.Name = "VistaPreviaFilasForm";
            this.Text = "Vista previa de filas a procesar";

            ((System.ComponentModel.ISupportInitialize)(this.dgvFilas)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
