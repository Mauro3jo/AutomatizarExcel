namespace Automatizacion_excel
{
    partial class Home
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnSeleccionarArchivo;
        private System.Windows.Forms.Label lblRutaArchivo;
        private System.Windows.Forms.FlowLayoutPanel panelBotones;
        private System.Windows.Forms.ProgressBar progressBar1;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.btnSeleccionarArchivo = new System.Windows.Forms.Button();
            this.lblRutaArchivo = new System.Windows.Forms.Label();
            this.panelBotones = new System.Windows.Forms.FlowLayoutPanel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();

            this.SuspendLayout();

            // btnSeleccionarArchivo
            this.btnSeleccionarArchivo.Location = new System.Drawing.Point(30, 20);
            this.btnSeleccionarArchivo.Size = new System.Drawing.Size(160, 40);
            this.btnSeleccionarArchivo.Text = "Seleccionar Excel";
            this.btnSeleccionarArchivo.Click += new System.EventHandler(this.btnSeleccionarArchivo_Click);

            // lblRutaArchivo
            this.lblRutaArchivo.Location = new System.Drawing.Point(210, 30);
            this.lblRutaArchivo.Size = new System.Drawing.Size(540, 20);
            this.lblRutaArchivo.Text = "Archivo no cargado";

            // panelBotones
            this.panelBotones.Location = new System.Drawing.Point(30, 80);
            this.panelBotones.Size = new System.Drawing.Size(740, 300); // Aumentado para más espacio
            this.panelBotones.AutoScroll = true;
            this.panelBotones.WrapContents = true;
            this.panelBotones.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;

            // progressBar1
            this.progressBar1.Location = new System.Drawing.Point(30, 390);
            this.progressBar1.Size = new System.Drawing.Size(740, 20);
            this.progressBar1.Visible = false;

            // Home
            this.ClientSize = new System.Drawing.Size(800, 440); // Aumentado también
            this.Controls.Add(this.btnSeleccionarArchivo);
            this.Controls.Add(this.lblRutaArchivo);
            this.Controls.Add(this.panelBotones);
            this.Controls.Add(this.progressBar1);
            this.Text = "Automatización Excel";
            this.ResumeLayout(false);
        }
    }
}
