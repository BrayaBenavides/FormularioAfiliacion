using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            dt.Columns.Add("Nombres");

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
    }
}
