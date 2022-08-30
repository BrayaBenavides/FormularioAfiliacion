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


            dt.Columns.Add("TrabajadorId");

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

            dt.DefaultView.RowFilter = $"TrabajadorId LIKE '{TxtFiltrar.Text}%'";
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

                    var Empleador = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[1].Value);
                    var Trabajador = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[1].Value);
                    var Conyuge = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[2].Value);
                    var Beneficiario1 = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    var Beneficiario2 = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    var Beneficiario3 = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    var Beneficiario4 = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    var Beneficiario5 = Convert.ToString(this.DataDetalles.SelectedRows[0].Cells[3].Value);
                    


                    File.WriteAllLines("campos.txt", campos);
                    string newFile = @"C:\Users\ticdesarrollo01\Downloads\" + Nombres + " " + Papellido + ".pdf";
                    pdfReader = new PdfReader(pdfTemplate);
                    PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    AcroFields pdfFormFields = pdfStamper.AcroFields;



                    //1. TIPO DE NOVEDAD

                    switch ( )  
                    {
                        case "1": // Afiliación
                            pdfFormFields.SetField("Casilla de verificación1", "0"); 
                            break;

                        case "2": // Novedad
                            pdfFormFields.SetField("Casilla de verificación2", "0"); 
                            break;

                        case "3": // Traslado
                            pdfFormFields.SetField("Casilla de verificación3", "0");
                            break;

                        default:
                            MessageBox.Show("Error de tipo de novedad afiliación");
                            break;
                    }

                    //Fecha de radicación
                    pdfFormFields.SetField("Texto1", " "); // Día 
                    pdfFormFields.SetField("Texto2", " "); // Mes 
                    pdfFormFields.SetField("Texto3", " "); // Año 


                    //3. DATOS DEL EMPLEADOR O ENTIDAD PAGADORA

                    switch ( ) // Tipo de identidad 
                    {
                        case "7": // Nit
                            pdfFormFields.SetField("Casilla de verificación16", "0");
                            break;

                        case "1": // Cedula de ciudadania (C.C.)
                            pdfFormFields.SetField("Casilla de verificación17", "0");
                            break;

                        case "3": // Cedula Extranjera (CE) 
                            pdfFormFields.SetField("Casilla de verificación21", "0");
                            break;

                        case "4": //  Pasaporte (PA)                                        
                            pdfFormFields.SetField("Casilla de verificación22", "0");
                            break;

                        case "2": // Carnet (C.D.)
                            pdfFormFields.SetField("Casilla de verificación23", "0");
                            break;

                        case "6": // Permiso Especial (P.E) 
                            pdfFormFields.SetField("Casilla de verificación24", "0");
                            break;

                        case "9": // Permiso Turismo (P.T) 
                            pdfFormFields.SetField("Casilla de verificación25", "0");
                            break;

                        default:
                            MessageBox.Show("Error en el tipo de documento, entidad pagadora");
                            break;
                    }

                    pdfFormFields.SetField("Texto7", " "); // No. de identidad pagadora
                    pdfFormFields.SetField("Texto9", " "); // Nombre o Razón Social

                    //Sector
                    switch ( ) 
                    {
                        case "1":  // Oficial Público
                            pdfFormFields.SetField("Casilla de verificación26", "0");
                            break;

                        case "2": // Privado
                            pdfFormFields.SetField("Casilla de verificación27", "0");
                            break;

                        case "3": // Mixto
                            pdfFormFields.SetField("Casilla de verificación28", "0");
                            break;

                        default:
                            MessageBox.Show("Error de sector, entidad pagadora");
                            break;
                    }

                    pdfFormFields.SetField("Texto10", " "); // Sucursal 
                    pdfFormFields.SetField("Texto13", " "); // Dirección
                    pdfFormFields.SetField("Texto12", " "); // Departamento
                    pdfFormFields.SetField("Texto11", " "); // Teléfono


                    // 5. DATOS BÁSICOS TRABAJADOR INDEPENDIENTE O PENSIONADO 

                    pdfFormFields.SetField("Texto16", " "); // Nombres 
                    pdfFormFields.SetField("Texto17", " "); // Primer apellido
                    pdfFormFields.SetField("Texto18", " "); // Segundo apellido

                    switch ( )
                    {
                        case "7": // Nit
                            pdfFormFields.SetField("Casilla de verificación45", "0");
                            break;

                        case "1": // Cedula de ciudadania (C.C.)
                            pdfFormFields.SetField("Casilla de verificación46", "0");
                            break;

                        case "3": //  Cedula Extranjera (CE)  
                            pdfFormFields.SetField("Casilla de verificación47", "0");
                            break;

                        case "4": // Pasaporte (PA) 
                            pdfFormFields.SetField("Casilla de verificación48", "0");
                            break;

                        case "2": // Carnet (C.D.)
                            pdfFormFields.SetField("Casilla de verificación49", "0");
                            break;

                        case "5": // Tarjeta de identidad (T.I)
                            pdfFormFields.SetField("Casilla de verificación50", "0");
                            break;

                        case "6": // Permiso Especial (P.E) 
                            pdfFormFields.SetField("Casilla de verificación51", "0");
                            break;

                        case "9": // Permiso Turismo (P.T) 
                            pdfFormFields.SetField("Casilla de verificación52", "0");
                            break;

                        default:
                            MessageBox.Show("Error en el tipo de documento, Trabajador");
                            break;
                    }

                    pdfFormFields.SetField("Texto19", " "); //No. documento

                    //Estado Civil
                    switch ( ) 
                    {
                        case "1": // Soltero (SO)
                            pdfFormFields.SetField("Casilla de verificación53", "0");
                            break;

                        case "2": // Casado(CA)
                            pdfFormFields.SetField("Casilla de verificación54", "0");
                            break;

                        case "3": // Separado (SE)  
                            pdfFormFields.SetField("Casilla de verificación55", "0");
                            break;

                        case "4": // Union Libre (UL) 
                            pdfFormFields.SetField("Casilla de verificación56", "0");
                            break;

                        case "5": // Viudo (VI)
                            pdfFormFields.SetField("Casilla de verificación57", "0");
                            break;

                        default:
                            MessageBox.Show("Error estado civil");
                            break;
                    }

                    //Fecha de nacimiento
                    pdfFormFields.SetField("Texto20", " "); // Año
                    pdfFormFields.SetField("Texto21", " "); // Mes 
                    pdfFormFields.SetField("Texto22", " "); // Día

                    //Género
                    switch ( ) 
                    {
                        case "2": // Masculino(M)
                            pdfFormFields.SetField("Casilla de verificación58", "0");
                            break;

                        case "1": // Femenino(F)
                            pdfFormFields.SetField("Casilla de verificación59", "0");
                            break;

                        case "3": // indefinido(I)  
                            pdfFormFields.SetField("Casilla de verificación60", "0");
                            break;

                        default:
                            MessageBox.Show("Error de género");
                            break;
                    }

                    //Nivel Ocupacional
                    switch ( ) 
                    {
                        case "1": // Operativo 
                            pdfFormFields.SetField("Casilla de verificación73", "0");
                            break;

                        case "2": // Administrativo
                            pdfFormFields.SetField("Casilla de verificación72", "0");
                            break;

                        case "3": // Directivo 
                            pdfFormFields.SetField("Casilla de verificación71", "0");
                            break;

                        case "4": // Staff
                            pdfFormFields.SetField("Casilla de verificación70", "0");
                            break;

                        case "5": // Ejecutivo
                            pdfFormFields.SetField("Casilla de verificación69", "0");
                            break;

                        case "6": // Profesional
                            pdfFormFields.SetField("Casilla de verificación68", "0");
                            break;

                        case "7": // Técnico
                            pdfFormFields.SetField("Casilla de verificación67", "0");
                            break;

                        default:
                            MessageBox.Show("Error, nivel ocupacional");
                            break;
                    }

                    //Nivel Educativo
                    switch () 
                    {
                        case "1": // Ninguno 
                            pdfFormFields.SetField("Casilla de verificación61", "0");
                        break;

                        case "2": // Primaria
                            pdfFormFields.SetField("Casilla de verificación62", "0");
                        break;

                        case "3": // Secundaria
                            pdfFormFields.SetField("Casilla de verificación63", "0");
                        break;

                        case "4": // Técnico
                            pdfFormFields.SetField("Casilla de verificación64", "0");
                        break;

                        case "5": // profesional
                            pdfFormFields.SetField("Casilla de verificación65", "0");
                        break;

                        case "6": // Otro
                            pdfFormFields.SetField("Casilla de verificación66", "0");
                        break;

                        default:
                            MessageBox.Show("Error, nivel ocupacional");
                        break;
                    }

                    //Ingreso a la empresa
                    pdfFormFields.SetField("Texto101", " "); // Año
                    pdfFormFields.SetField("Texto24", " "); // Mes 
                    pdfFormFields.SetField("Texto25", " "); // Día

                    
                    pdfFormFields.SetField("Texto26", " "); // Horas/mes

                    //Trabajador
                    switch () 
                    {
                        case "1": // urbano(UR) 
                            pdfFormFields.SetField("Casilla de verificación74", "0");
                            break;

                        case "2": // Rural(Ru)
                            pdfFormFields.SetField("Casilla de verificación75", "0");
                            break;

                        default:
                            MessageBox.Show("Error, Trabajador");
                            break;
                    }

                    //Salario básico/mesada
                    pdfFormFields.SetField("Texto27", " ");

                    //Celular
                    pdfFormFields.SetField("Texto28", " ");

                    //EPS (Solo para independientes)
                    pdfFormFields.SetField("Texto29", " ");

                    //AFP (Solo para independientes)
                    pdfFormFields.SetField("Texto30", " ");

                    //Dirección vivienda
                    pdfFormFields.SetField("Texto100", " ");

                    //Municipio
                    pdfFormFields.SetField("Texto35", " ");

                    //Departamento
                    pdfFormFields.SetField("Texto32", " ");

                    //Zona vivienda
                    switch () 
                    {
                        case "1": // urbano(UR) 
                            pdfFormFields.SetField("Casilla de verificación77", "0");
                            break;

                        case "2": // Rural(Ru)
                            pdfFormFields.SetField("Casilla de verificación76", "0");
                            break;

                        default:
                            MessageBox.Show("Error, zona vivienda");
                            break;
                    }

                    //Teléfono vivienda
                    pdfFormFields.SetField("Texto31", " ");

                    //Dirección trabajo
                    pdfFormFields.SetField("Texto38", " ");

                    //Municipio
                    pdfFormFields.SetField("Texto36", " ");

                    //Departamento
                    pdfFormFields.SetField("Texto102", " ");

                    //Zona Trabajo
                    switch () 
                    {
                        case "1": // urbano(UR) 
                            pdfFormFields.SetField("Casilla de verificación78", "0");
                            break;

                        case "2": // Rural(Ru)
                            pdfFormFields.SetField("Casilla de verificación79", "0");
                            break;

                        default:
                            MessageBox.Show("Error, zona trabajo");
                            break;
                    }

                    //Teléfono trabajo 
                    pdfFormFields.SetField("Texto33", " ");

                    //Correo Electrónico
                    pdfFormFields.SetField("Texto39", " ");

                    //País y Ciudad de Residencia (Solo para Residentes en el Exterior)
                    pdfFormFields.SetField("Texto40", " ");

                    //Programa (Sólo no dependientes)
                    pdfFormFields.SetField("Texto41", " ");


                    //7. Información del cónyuge o compañer@

                    switch ( ) 
                    {
                        case "1": // Cedula de ciudadania (C.C.)
                            pdfFormFields.SetField("Casilla de verificación82", "0");
                            break;

                        case "3": //  Cedula Extranjera (CE)  
                            pdfFormFields.SetField("Casilla de verificación83", "0");
                            break;

                        case "4": // Pasaporte (PA) 
                            pdfFormFields.SetField("Casilla de verificación84", "0");
                            break;

                        case "2": // Carnet (C.D.)
                            pdfFormFields.SetField("Casilla de verificación85", "0");
                            break;

                        case "5": // Tarjeta de identidad (T.I)
                            pdfFormFields.SetField("Casilla de verificación86", "0");
                            break;

                        case "6": // Permiso Especial (P.E) 
                            pdfFormFields.SetField("Casilla de verificación87", "0");
                            break;

                        case "9": // Permiso Turismo (P.T) 
                            pdfFormFields.SetField("Casilla de verificación88", "0");
                            break;

                        default:
                            MessageBox.Show("Error en el tipo de documento, Conyuge");
                            break;
                    }

                    pdfFormFields.SetField("Texto47", ""); //No. documento Conyuge

                    //Fecha de nacimiento
                    pdfFormFields.SetField("Texto48", " "); // Año
                    pdfFormFields.SetField("Texto49", " "); // Mes 
                    pdfFormFields.SetField("Texto50", " "); // Día

                    //Género Conyuge
                    switch ( ) 
                    {
                        case "2": // Masculino(M)
                            pdfFormFields.SetField("Casilla de verificación89", "0");
                            break;

                        case "1": // Femenino(F)
                            pdfFormFields.SetField("Casilla de verificación90", "0");
                            break;

                        case "3": // indefinido(I)  
                            pdfFormFields.SetField("Casilla de verificación91", "0");
                            break;

                        default:
                            MessageBox.Show("Error de género, Conyuge");
                            break;
                    }

                    //Trabaja Conyuge
                    switch ( ) 
                    {
                        case "SI": // Si
                            pdfFormFields.SetField("Casilla de verificación92", "0");
                            break;

                        case "NO": // No
                            pdfFormFields.SetField("Casilla de verificación93", "0");
                            break;

                        case "NULL":  
                            
                            break;

                        default:
                            MessageBox.Show("Error de trabaja, Conyuge");
                            break;
                    }

                    //salario básico conyuge 
                    pdfFormFields.SetField("Texto51", " ");

                    //Fecha de Ingreso a la empresa
                    pdfFormFields.SetField("Texto52", " "); // Año
                    pdfFormFields.SetField("Texto53", " "); // Mes 
                    pdfFormFields.SetField("Texto54", " "); // Día

                    pdfFormFields.SetField("Texto57", " "); // Nombres 
                    pdfFormFields.SetField("Texto56", " "); // Primer apellido
                    pdfFormFields.SetField("Texto55", " "); // Segundo apellido

                    pdfFormFields.SetField("Texto58", " "); // Razón social
                    pdfFormFields.SetField("Texto59", " "); // NIT


                    switch ( ) // Recibe subsidio
                    {
                        case "SI": // Si
                            pdfFormFields.SetField("Casilla de verificación94", "0");
                            break;

                        case "NO": // No
                            pdfFormFields.SetField("Casilla de verificación95", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de Recibe subsidio");
                            break;
                    }

                    pdfFormFields.SetField("Texto60", " "); // Caja por la cual recibe subsidio


                    //8. INFORMACIÓN GRUPO FAMILIAR 


                    // 1. Tipo 
                    pdfFormFields.SetField("Texto74", " ");

                    //No. Documento de identificación
                    pdfFormFields.SetField("Texto65", " ");

                    //Nombres
                    pdfFormFields.SetField("Texto34", " ");

                    //Apellidos
                    pdfFormFields.SetField("Texto74", " ");

                    switch ( ) //Parentesco
                    {
                        case "1": // Hijo 
                            pdfFormFields.SetField("Casilla de verificación96", "0");
                            break;

                        case "3": // padre 
                            pdfFormFields.SetField("Casilla de verificación97", "0");
                            break;

                        case "4": // Hermano
                            pdfFormFields.SetField("Casilla de verificación98", "0");
                            break;

                        case "2": // Hijastro 
                            pdfFormFields.SetField("Casilla de verificación99", "0");
                            break;

                        case "5": // Custodia
                            pdfFormFields.SetField("Casilla de verificación100", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de parentesco");
                            break;
                    }

                    //Fecha de Nacimiento 
                    pdfFormFields.SetField("Texto91", " "); // Año
                    pdfFormFields.SetField("Texto93", " "); // Mes 
                    pdfFormFields.SetField("Texto92", " "); // Día

                    // 2. Tipo 
                    pdfFormFields.SetField("Texto75", " ");

                    //No. Documento de identificación
                    pdfFormFields.SetField("Texto63", " ");

                    //Nombres
                    pdfFormFields.SetField("Texto64", " ");

                    //Apellidos
                    pdfFormFields.SetField("Texto62", " ");

                    switch ( ) //Parentesco
                    {
                        case "1": // Hijo 
                            pdfFormFields.SetField("Casilla de verificación105", "0");
                            break;

                        case "3": // padre 
                            pdfFormFields.SetField("Casilla de verificación104", "0");
                            break;

                        case "4": // Hermano
                            pdfFormFields.SetField("Casilla de verificación103", "0");
                            break;

                        case "2": // Hijastro 
                            pdfFormFields.SetField("Casilla de verificación102", "0");
                            break;

                        case "5": // Custodia
                            pdfFormFields.SetField("Casilla de verificación101", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de parentesco");
                            break;
                    }

                    //Fecha de Nacimiento 
                    pdfFormFields.SetField("Texto88", " "); // Año
                    pdfFormFields.SetField("Texto90", " "); // Mes 
                    pdfFormFields.SetField("Texto89", " "); // Día

                    // 3. Tipo 
                    pdfFormFields.SetField("Texto76", " ");

                    //No. Documento de identificación
                    pdfFormFields.SetField("Texto37", " ");

                    //Nombres
                    pdfFormFields.SetField("Texto66", " ");

                    //Apellidos
                    pdfFormFields.SetField("Texto69", " ");

                    switch ( ) //Parentesco
                    {
                        case "1": // Hijo 
                            pdfFormFields.SetField("Casilla de verificación110", "0");
                            break;

                        case "3": // padre 
                            pdfFormFields.SetField("Casilla de verificación109", "0");
                            break;

                        case "4": // Hermano
                            pdfFormFields.SetField("Casilla de verificación108", "0");
                            break;

                        case "2": // Hijastro 
                            pdfFormFields.SetField("Casilla de verificación107", "0");
                            break;

                        case "5": // Custodia
                            pdfFormFields.SetField("Casilla de verificación106", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de parentesco");
                            break;
                    }

                    //Fecha de Nacimiento 
                    pdfFormFields.SetField("Texto85", " "); // Año
                    pdfFormFields.SetField("Texto87", " "); // Mes 
                    pdfFormFields.SetField("Texto86", " "); // Día

                    // 4. Tipo 
                    pdfFormFields.SetField("Texto77", " ");

                    //No. Documento de identificación
                    pdfFormFields.SetField("Texto67", " ");

                    //Nombres
                    pdfFormFields.SetField("Texto70", " ");

                    //Apellidos
                    pdfFormFields.SetField("Texto68", " ");

                    switch ( ) //Parentesco
                    {
                        case "1": // Hijo 
                            pdfFormFields.SetField("Casilla de verificación115", "0");
                            break;

                        case "3": // padre 
                            pdfFormFields.SetField("Casilla de verificación114", "0");
                            break;

                        case "4": // Hermano
                            pdfFormFields.SetField("Casilla de verificación113", "0");
                            break;

                        case "2": // Hijastro 
                            pdfFormFields.SetField("Casilla de verificación112", "0");
                            break;

                        case "5": // Custodia
                            pdfFormFields.SetField("Casilla de verificación111", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de parentesco");
                            break;
                    }

                    //Fecha de Nacimiento 
                    pdfFormFields.SetField("Texto82", " "); // Año
                    pdfFormFields.SetField("Texto84", " "); // Mes 
                    pdfFormFields.SetField("Texto83", " "); // Día

                    // 5. Tipo 
                    pdfFormFields.SetField("Texto78", " ");

                    //No. Documento de identificación
                    pdfFormFields.SetField("Texto71", " ");

                    //Nombres
                    pdfFormFields.SetField("Texto72", " ");

                    //Apellidos
                    pdfFormFields.SetField("Texto73", " ");

                    switch ( ) //Parentesco
                    {
                        case "1": // Hijo 
                            pdfFormFields.SetField("Casilla de verificación120", "0");
                            break;

                        case "3": // padre 
                            pdfFormFields.SetField("Casilla de verificación119", "0");
                            break;

                        case "4": // Hermano
                            pdfFormFields.SetField("Casilla de verificación118", "0");
                            break;

                        case "2": // Hijastro 
                            pdfFormFields.SetField("Casilla de verificación117", "0");
                            break;

                        case "5": // Custodia
                            pdfFormFields.SetField("Casilla de verificación116", "0");
                            break;

                        case "NULL":

                            break;

                        default:
                            MessageBox.Show("Error de parentesco");
                            break;
                    }

                    //Fecha de Nacimiento 
                    pdfFormFields.SetField("Texto79", " "); // Año
                    pdfFormFields.SetField("Texto80", " "); // Mes 
                    pdfFormFields.SetField("Texto81", " "); // Día



                    //AUTORIZACIÓN PARA LA UTILIZACIÓN DE DATOS PERSONALES

                    pdfFormFields.SetField("Casilla de verificación121", "0"); // Año














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






 
