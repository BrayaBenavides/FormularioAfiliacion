using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

namespace FormularioExcel
{
    public partial class Form1 : Form
    {
        
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();

        public Form1()
        {
            InitializeComponent();
            

            dt.Columns.Add("Id");

        }

        private void BtnImportar_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog
            {
                Filter = "Excel | *.csv; *.xlsx;",
                Title = "Seleccionar archivo"

            };

            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                DataDetalles.DataSource = ImportarDatos(OpenFileDialog.FileName);
            }

            LblFile.Text = OpenFileDialog.FileName;
        }
        

        DataView ImportarDatos(string nombrearchivo)
        {
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended properties = 'Excel 12.0;'", nombrearchivo);

            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            OleDbCommand consulta = new OleDbCommand("select * from [Hoja1$]", conector);

            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta
            };

            adaptador.Fill(dt);
            adaptador.Fill(ds);
            conector.Close();

            return ds.Tables[0].DefaultView;
        }

        private void TxtFiltrar_TextChanged(object sender, EventArgs e)
        {
            
            dt.DefaultView.RowFilter = $"Id LIKE '{TxtFiltrar.Text}%'";
            DataDetalles.DataSource = dt;
        }

        private void DataDetalles_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
           
            try
            {
                if (MessageBox.Show("Exportar a PDF?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                { 
                    for (int i = 1; i <= 3; i++)
                    {
                        var prueba = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[i].Value);

                        PdfDocument pdfDocument = new PdfDocument(new PdfWriter(new FileStream("C:/Users/ticdesarrollo01/source/repos/FormularioAfiliacion/bin/Debug/Prueba.pdf", FileMode.Create, FileAccess.Write)));
                        Document document = new Document(pdfDocument);

                        document.Add(new Paragraph(prueba));

                        document.Close();
                    }
                }
                else
                {
                    // user clicked no

                }
            }
            catch (Exception)
            {
                MessageBox.Show("Seleccione toda la fila?");
            }
        }
    }
}
