using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;


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
                Filter = "Excel | *.xlsx;",
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


        private void BtnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Exportar a PDF?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {

                    string pdfTemplate = @"C:\Users\ticdesarrollo01\source\repos\FormularioAfiliacion\Formulario.pdf";
                    PdfReader pdfReader = new PdfReader(pdfTemplate);
                    AcroFields af = pdfReader.AcroFields;
                    List<string> campos = new List<string>();
                    foreach (KeyValuePair<string, AcroFields.Item> kvp in af.Fields)
                    {
                        string fieldName = kvp.Key.ToString();
                        string fieldValue = af.GetField(kvp.Key.ToString());
                        campos.Add(fieldName + " " + fieldValue);
                    }

                    var Nombres = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[1].Value);
                    var Papellido = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[2].Value);
                    var Sapellido = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    var Tdocumento = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[4].Value);
                    var NoIdentidad = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[5].Value);


                    File.WriteAllLines("campos.txt", campos);
                    string newFile = @"C:\Users\ticdesarrollo01\Downloads\" + Nombres + " " + Papellido + ".pdf";
                    pdfReader = new PdfReader(pdfTemplate);
                    PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    AcroFields pdfFormFields = pdfStamper.AcroFields;

                    pdfFormFields.SetField("Texto16", Nombres);
                    pdfFormFields.SetField("Texto17", Papellido);
                    pdfFormFields.SetField("Texto18", Sapellido);

                    switch (Tdocumento)
                    {
                        case "NIT":
                            pdfFormFields.SetField("Casilla de verificación45", "0");
                            break;

                        case "C.C.":
                            pdfFormFields.SetField("Casilla de verificación46", "0");
                            break;

                        case "C.E.":
                            pdfFormFields.SetField("Casilla de verificación47", "0");
                            break;

                        case "PA":
                            pdfFormFields.SetField("Casilla de verificación48", "0");
                            break;

                        case "C.D.":
                            pdfFormFields.SetField("Casilla de verificación49", "0");
                            break;

                        case "T.I.":
                            pdfFormFields.SetField("Casilla de verificación50", "0");
                            break;

                        case "PE":
                            pdfFormFields.SetField("Casilla de verificación51", "0");
                            break;

                        case "PT":
                            pdfFormFields.SetField("Casilla de verificación52", "0");
                            break;

                        default:
                            MessageBox.Show("Error en el tipo de documento");
                            break;
                    }

                    pdfFormFields.SetField("Texto19", NoIdentidad);

                    pdfStamper.FormFlattening = true;
                    pdfStamper.Close();

                    Process.Start(newFile);
                }
                else
                {
                    MessageBox.Show("Error");
                }

            }
            catch (Exception)
            {
                MessageBox.Show("Seleccione toda la fila");
            }

        }
    }
}






 
