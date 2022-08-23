
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
            this.TxtFiltrar = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.DataDetalles)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnImportar
            // 
            this.BtnImportar.Location = new System.Drawing.Point(53, 52);
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
            this.LblFile.Location = new System.Drawing.Point(155, 57);
            this.LblFile.Name = "LblFile";
            this.LblFile.Size = new System.Drawing.Size(35, 13);
            this.LblFile.TabIndex = 1;
            this.LblFile.Text = "label1";
            // 
            // DataDetalles
            // 
            this.DataDetalles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataDetalles.Location = new System.Drawing.Point(53, 130);
            this.DataDetalles.Name = "DataDetalles";
            this.DataDetalles.Size = new System.Drawing.Size(684, 264);
            this.DataDetalles.TabIndex = 2;
            // 
            // TxtFiltrar
            // 
            this.TxtFiltrar.Location = new System.Drawing.Point(637, 90);
            this.TxtFiltrar.Name = "TxtFiltrar";
            this.TxtFiltrar.Size = new System.Drawing.Size(100, 20);
            this.TxtFiltrar.TabIndex = 3;
            this.TxtFiltrar.TextChanged += new System.EventHandler(this.TxtFiltrar_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(599, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Filtrar";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TxtFiltrar);
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
        private System.Windows.Forms.TextBox TxtFiltrar;
        private System.Windows.Forms.Label label1;
    }
}

