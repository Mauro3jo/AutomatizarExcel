namespace Automatizacion_excel
{
    partial class Home
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button btnSeleccionarArchivo;
        private System.Windows.Forms.Label lblRutaArchivo;
        private System.Windows.Forms.Button btnVerTasas;
        private System.Windows.Forms.Button btnVerIIBB;
        private System.Windows.Forms.FlowLayoutPanel panelBotones;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Panel panelLinea1;
        private System.Windows.Forms.Button btnFiserv;
        private System.Windows.Forms.Button btnQR;

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
            this.btnVerTasas = new System.Windows.Forms.Button();
            this.btnVerIIBB = new System.Windows.Forms.Button();
            this.panelBotones = new System.Windows.Forms.FlowLayoutPanel();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panelLinea1 = new System.Windows.Forms.Panel();
            this.btnFiserv = new System.Windows.Forms.Button();
            this.btnQR = new System.Windows.Forms.Button();

            this.SuspendLayout();

            // 
            // btnFiserv
            // 
            this.btnFiserv.Location = new System.Drawing.Point(30, 60);
            this.btnFiserv.Name = "btnFiserv";
            this.btnFiserv.Size = new System.Drawing.Size(160, 30);
            this.btnFiserv.TabIndex = 7;
            this.btnFiserv.Text = "Fiserv";
            this.btnFiserv.UseVisualStyleBackColor = true;
            this.btnFiserv.Click += new System.EventHandler(this.btnFiserv_Click);

            // 
            // btnQR
            // 
            this.btnQR.Location = new System.Drawing.Point(200, 60);
            this.btnQR.Name = "btnQR";
            this.btnQR.Size = new System.Drawing.Size(160, 30);
            this.btnQR.TabIndex = 8;
            this.btnQR.Text = "QR";
            this.btnQR.UseVisualStyleBackColor = true;
            this.btnQR.Click += new System.EventHandler(this.btnQR_Click);

            // 
            // btnSeleccionarArchivo
            // 
            this.btnSeleccionarArchivo.Location = new System.Drawing.Point(30, 20);
            this.btnSeleccionarArchivo.Name = "btnSeleccionarArchivo";
            this.btnSeleccionarArchivo.Size = new System.Drawing.Size(160, 40);
            this.btnSeleccionarArchivo.TabIndex = 0;
            this.btnSeleccionarArchivo.Text = "Seleccionar Excel";
            this.btnSeleccionarArchivo.UseVisualStyleBackColor = true;
            this.btnSeleccionarArchivo.Click += new System.EventHandler(this.btnSeleccionarArchivo_Click);

            // 
            // lblRutaArchivo
            // 
            this.lblRutaArchivo.Location = new System.Drawing.Point(210, 30);
            this.lblRutaArchivo.Name = "lblRutaArchivo";
            this.lblRutaArchivo.Size = new System.Drawing.Size(480, 20);
            this.lblRutaArchivo.TabIndex = 1;
            this.lblRutaArchivo.Text = "Archivo no cargado";

            // 
            // btnVerTasas
            // 
            this.btnVerTasas.Location = new System.Drawing.Point(700, 20);
            this.btnVerTasas.Name = "btnVerTasas";
            this.btnVerTasas.Size = new System.Drawing.Size(180, 30);
            this.btnVerTasas.TabIndex = 2;
            this.btnVerTasas.Text = "📊 Ver y editar tasas";
            this.btnVerTasas.UseVisualStyleBackColor = true;
            this.btnVerTasas.Click += new System.EventHandler(this.btnVerTasas_Click);

            // 
            // btnVerIIBB
            // 
            this.btnVerIIBB.Location = new System.Drawing.Point(700, 55);
            this.btnVerIIBB.Name = "btnVerIIBB";
            this.btnVerIIBB.Size = new System.Drawing.Size(180, 30);
            this.btnVerIIBB.TabIndex = 3;
            this.btnVerIIBB.Text = "📋 Ver IIBB por provincia";
            this.btnVerIIBB.UseVisualStyleBackColor = true;
            this.btnVerIIBB.Click += new System.EventHandler(this.btnVerIIBB_Click);

            // 
            // panelLinea1
            // 
            this.panelLinea1.BackColor = System.Drawing.Color.DarkGray;
            this.panelLinea1.Location = new System.Drawing.Point(30, 100);
            this.panelLinea1.Name = "panelLinea1";
            this.panelLinea1.Size = new System.Drawing.Size(850, 2);
            this.panelLinea1.TabIndex = 6;

            // 
            // panelBotones
            // 
            this.panelBotones.Location = new System.Drawing.Point(30, 110);
            this.panelBotones.Name = "panelBotones";
            this.panelBotones.Size = new System.Drawing.Size(880, 320);
            this.panelBotones.TabIndex = 4;
            this.panelBotones.AutoScroll = true;
            this.panelBotones.WrapContents = true;
            this.panelBotones.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;

            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(30, 440);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(820, 20);
            this.progressBar1.TabIndex = 5;
            this.progressBar1.Visible = false;

            // 
            // Home
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(940, 480);
            this.Controls.Add(this.btnFiserv);
            this.Controls.Add(this.btnQR);
            this.Controls.Add(this.btnSeleccionarArchivo);
            this.Controls.Add(this.lblRutaArchivo);
            this.Controls.Add(this.btnVerTasas);
            this.Controls.Add(this.btnVerIIBB);
            this.Controls.Add(this.panelLinea1);
            this.Controls.Add(this.panelBotones);
            this.Controls.Add(this.progressBar1);
            this.Name = "Home";
            this.Text = "Automatización Excel";
            this.ResumeLayout(false);
        }
    }
}
