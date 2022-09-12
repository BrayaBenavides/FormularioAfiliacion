
namespace FormularioExcel
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.BtnImportar = new System.Windows.Forms.Button();
            this.LblFile = new System.Windows.Forms.Label();
            this.DataDetalles = new System.Windows.Forms.DataGridView();
            this.BtnExportar = new System.Windows.Forms.Button();
            this.LblPDF = new System.Windows.Forms.Label();
            this.LblError = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.Lblguardar = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DataDetalles)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnImportar
            // 
            this.BtnImportar.Location = new System.Drawing.Point(53, 47);
            this.BtnImportar.Name = "BtnImportar";
            this.BtnImportar.Size = new System.Drawing.Size(75, 23);
            this.BtnImportar.TabIndex = 0;
            this.BtnImportar.Text = "Importar";
            this.BtnImportar.UseVisualStyleBackColor = true;
            this.BtnImportar.Click += new System.EventHandler(this.BtnImportar_Click);
            // 
            // LblFile
            // 
            this.LblFile.AutoSize = true;
            this.LblFile.Location = new System.Drawing.Point(134, 52);
            this.LblFile.Name = "LblFile";
            this.LblFile.Size = new System.Drawing.Size(33, 13);
            this.LblFile.TabIndex = 1;
            this.LblFile.Text = "Ruta:";
            // 
            // DataDetalles
            // 
            this.DataDetalles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataDetalles.Location = new System.Drawing.Point(53, 169);
            this.DataDetalles.Name = "DataDetalles";
            this.DataDetalles.Size = new System.Drawing.Size(684, 202);
            this.DataDetalles.TabIndex = 2;
            // 
            // BtnExportar
            // 
            this.BtnExportar.Location = new System.Drawing.Point(646, 391);
            this.BtnExportar.Name = "BtnExportar";
            this.BtnExportar.Size = new System.Drawing.Size(91, 23);
            this.BtnExportar.TabIndex = 7;
            this.BtnExportar.Text = "Exportar PDF";
            this.BtnExportar.UseVisualStyleBackColor = true;
            this.BtnExportar.Click += new System.EventHandler(this.BtnExportar_Click);
            // 
            // LblPDF
            // 
            this.LblPDF.AutoSize = true;
            this.LblPDF.Location = new System.Drawing.Point(344, 105);
            this.LblPDF.Name = "LblPDF";
            this.LblPDF.Size = new System.Drawing.Size(28, 13);
            this.LblPDF.TabIndex = 8;
            this.LblPDF.Text = "PDF";
            this.LblPDF.Visible = false;
            // 
            // LblError
            // 
            this.LblError.AutoSize = true;
            this.LblError.Location = new System.Drawing.Point(60, 153);
            this.LblError.Name = "LblError";
            this.LblError.Size = new System.Drawing.Size(40, 13);
            this.LblError.TabIndex = 9;
            this.LblError.Text = "Errores";
            this.LblError.Visible = false;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(53, 377);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(188, 13);
            this.progressBar.TabIndex = 10;
            this.progressBar.Visible = false;
            // 
            // Lblguardar
            // 
            this.Lblguardar.AutoSize = true;
            this.Lblguardar.Location = new System.Drawing.Point(527, 401);
            this.Lblguardar.Name = "Lblguardar";
            this.Lblguardar.Size = new System.Drawing.Size(35, 13);
            this.Lblguardar.TabIndex = 11;
            this.Lblguardar.Text = "label1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Lblguardar);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.LblError);
            this.Controls.Add(this.LblPDF);
            this.Controls.Add(this.BtnExportar);
            this.Controls.Add(this.DataDetalles);
            this.Controls.Add(this.LblFile);
            this.Controls.Add(this.BtnImportar);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.DataDetalles)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnImportar;
        private System.Windows.Forms.Label LblFile;
        private System.Windows.Forms.DataGridView DataDetalles;
        private System.Windows.Forms.Button BtnExportar;
        private System.Windows.Forms.Label LblPDF;
        private System.Windows.Forms.Label LblError;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label Lblguardar;
    }
}

