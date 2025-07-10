namespace Automatizacion_excel.Formularios
{
    partial class FormIIBB
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.DataGridView dgvIIBB;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.dgvIIBB = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgvIIBB)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvIIBB
            // 
            this.dgvIIBB.AllowUserToAddRows = false;
            this.dgvIIBB.AllowUserToDeleteRows = false;
            this.dgvIIBB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvIIBB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvIIBB.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvIIBB.Location = new System.Drawing.Point(0, 0);
            this.dgvIIBB.Name = "dgvIIBB";
            this.dgvIIBB.Size = new System.Drawing.Size(600, 400);
            this.dgvIIBB.TabIndex = 0;
            // 
            // FormIIBB
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(600, 400);
            this.Controls.Add(this.dgvIIBB);
            this.Name = "FormIIBB";
            this.Text = "📋 Ver IIBB por provincia";
            ((System.ComponentModel.ISupportInitialize)(this.dgvIIBB)).EndInit();
            this.ResumeLayout(false);
        }
    }
}
