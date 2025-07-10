namespace Automatizacion_excel.Formularios
{
    partial class FormTasas
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView dgvTasas;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        private void InitializeComponent()
        {
            this.dgvTasas = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTasas)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvTasas
            // 
            this.dgvTasas.AllowUserToAddRows = false;
            this.dgvTasas.AllowUserToDeleteRows = false;
            this.dgvTasas.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvTasas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTasas.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvTasas.Location = new System.Drawing.Point(0, 0);
            this.dgvTasas.Name = "dgvTasas";
            this.dgvTasas.Size = new System.Drawing.Size(800, 450);
            this.dgvTasas.TabIndex = 0;
            // 
            // FormTasas
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dgvTasas);
            this.Name = "FormTasas";
            this.Text = "📊 Ver y editar tasas";
            ((System.ComponentModel.ISupportInitialize)(this.dgvTasas)).EndInit();
            this.ResumeLayout(false);
        }

        #endregion
    }
}
