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

                   
                    int trabajador = 0;
                    int conyuge = 0;
                    int Benefi = 0;
                    string[] DatosTrabajador = new string[39];
                    string[] DatosConyuge = new string[39];
                    string[] DatosBeneficiario = new string[39];
                    string[][] ArrayConyuge = new string[this.DataDetalles.Rows.Count][];
                    string[][] ArrayTrabajador = new string[this.DataDetalles.Rows.Count][];
                    string[][] ArrayBeneficiario = new string[this.DataDetalles.Rows.Count][];
                    


                    foreach (DataGridViewRow Detalles in this.DataDetalles.Rows)
                    {
                        if(Convert.ToString(Detalles.Cells[4].Value) == "0")
                        {
                            for (int i = 1; i <= 38; i++)
                            {

                                DatosTrabajador[i] = Convert.ToString(Detalles.Cells[i].Value);
                            }
                            ArrayTrabajador[trabajador] = DatosTrabajador;
                            DatosTrabajador = new string[39];
                            trabajador++;
                        }
                        else if (Convert.ToString(Detalles.Cells[4].Value) == "1")
                        {
                            for (int i = 6; i <= 34; i++)
                            {
                                DatosConyuge[i-5] = Convert.ToString(Detalles.Cells[i].Value);
                            }
                            ArrayConyuge[conyuge] = DatosConyuge;
                            DatosConyuge = new string[39];
                            conyuge++;

                        }
                        else if(Convert.ToString(Detalles.Cells[4].Value) == "2")
                        {
                            for (int i = 6; i <= 32; i++)
                            {
                                DatosBeneficiario[i-5] = Convert.ToString(Detalles.Cells[i].Value);
                            }
                            ArrayBeneficiario[Benefi] = DatosBeneficiario;
                            DatosBeneficiario = new string[39];
                            Benefi++;
                        }
                    }

                   
                    foreach (var Trabajador in ArrayTrabajador)
                    {
                        if (Trabajador == null)
                        {
                            break;
                        }
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
                        File.WriteAllLines("campos.txt", campos);
                        string newFile = @"C:\Users\ticdesarrollo01\Downloads\" + Trabajador[10] + " " + Trabajador[12] + ".pdf";
                        pdfReader = new PdfReader(pdfTemplate);
                        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                        AcroFields pdfFormFields = pdfStamper.AcroFields;

                        //1. TIPO DE NOVEDAD

                        switch ("")
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

                        }

                        //Fecha de radicación
                        pdfFormFields.SetField("Texto1", Trabajador[5].Substring(3, 2)); // Día 
                        pdfFormFields.SetField("Texto2", Trabajador[5].Substring(0, 2)); // Mes 
                        pdfFormFields.SetField("Texto3", Trabajador[5].Substring(6, 4)); // Año


                        //3. DATOS DEL EMPLEADOR O ENTIDAD PAGADORA

                        switch (Trabajador[1]) // Tipo de identidad 
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

                           
                        }

                            pdfFormFields.SetField("Texto7", Trabajador[2] + "-" + Trabajador[3]); // No. de identidad pagadora
                            pdfFormFields.SetField("Texto9", " "); // Nombre o Razón Social


                            //Sector
                            switch (" ")
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

                                
                            }

                            pdfFormFields.SetField("Texto10", " "); // Sucursal 
                            pdfFormFields.SetField("Texto13", " "); // Dirección
                            pdfFormFields.SetField("Texto12", " "); // Departamento
                            pdfFormFields.SetField("Texto11", " "); // Teléfono



                            // 5. DATOS BÁSICOS TRABAJADOR INDEPENDIENTE O PENSIONADO

                            pdfFormFields.SetField("Texto16", Trabajador[10] + " " + Trabajador[11]); // Nombres 
                            pdfFormFields.SetField("Texto17", Trabajador[12]); // Primer apellido
                            pdfFormFields.SetField("Texto18", Trabajador[13]); // Segundo apellido


                            switch (Trabajador[7])
                            {
                                case "7": // Nit
                                    pdfFormFields.SetField("Casilla de verificación45", "0");
                                    break;

                                case "1": //Cédula de ciudadania
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

                                case "PT":
                                    pdfFormFields.SetField("Casilla de verificación52", "0");
                                    break;

                                
                            }

                            pdfFormFields.SetField("Texto19", Trabajador[8]); //No. documento

                            //Estado Civil
                            switch (Trabajador[17])
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

                                
                            }

                            //Fecha de nacimiento
                            pdfFormFields.SetField("Texto20", Trabajador[14].Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto21", Trabajador[14].Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto22", Trabajador[14].Substring(3, 2)); // Día

                            //Género
                            switch (Trabajador[15])
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

                               
                            }

                            //Nivel Ocupacional
                            switch (Trabajador[18])
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
                            }

                            //Nivel Educativo
                            switch (Trabajador[18])
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

                               
                            }

                            //Ingreso a la empresa
                            pdfFormFields.SetField("Texto101", Trabajador[19].Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto24", Trabajador[19].Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto25", Trabajador[19].Substring(3, 2)); // Día


                            pdfFormFields.SetField("Texto26", Trabajador[20]); // Horas/mes

                            //Trabajador
                            switch (" ")
                            {
                                case "1": // urbano(UR) 
                                    pdfFormFields.SetField("Casilla de verificación74", "0");
                                    break;

                                case "2": // Rural(Ru)
                                    pdfFormFields.SetField("Casilla de verificación75", "0");
                                    break;

                                
                            }

                            //Salario básico/mesada
                            pdfFormFields.SetField("Texto27", Trabajador[21]);

                            //Celular
                            pdfFormFields.SetField("Texto28", Trabajador[23]);

                            //EPS (Solo para independientes)
                            pdfFormFields.SetField("Texto29", " ");

                            //AFP (Solo para independientes)
                            pdfFormFields.SetField("Texto30", " ");

                            //Dirección vivienda
                            pdfFormFields.SetField("Texto100", Trabajador[24]);

                            //Municipio
                            switch (Trabajador[26])
                            {
                                case "11001":
                                    pdfFormFields.SetField("Texto35", "BOGOTA D.C");
                                    break;
                              
                            }

                            //Departamento
                            switch (Trabajador[25])
                            {
                                case "11":
                                    pdfFormFields.SetField("Texto32", "BOGOTA");
                                    break;
                            }


                            //Zona vivienda
                            switch (" ")
                            {
                                case "1": // urbano(UR) 
                                    pdfFormFields.SetField("Casilla de verificación77", "0");
                                    break;

                                case "2": // Rural(Ru)
                                    pdfFormFields.SetField("Casilla de verificación76", "0");
                                    break;                              
                            }

                            //Teléfono vivienda
                            pdfFormFields.SetField("Texto31", Trabajador[22]);

                            //Dirección trabajo
                            pdfFormFields.SetField("Texto38", " ");

                            //Municipio
                            pdfFormFields.SetField("Texto36", " ");

                            //Departamento
                            pdfFormFields.SetField("Texto102", " ");

                            //Zona Trabajo
                            switch (" ")
                            {
                                case "1": // urbano(UR) 
                                    pdfFormFields.SetField("Casilla de verificación78", "0");
                                    break;

                                case "2": // Rural(Ru)
                                    pdfFormFields.SetField("Casilla de verificación79", "0");
                                    break;

                             
                            }

                            //Teléfono trabajo 
                            pdfFormFields.SetField("Texto33", " ");

                            //Correo Electrónico
                            pdfFormFields.SetField("Texto39", Trabajador[27]);

                            //País y Ciudad de Residencia (Solo para Residentes en el Exterior)
                            pdfFormFields.SetField("Texto40", " ");

                            //Programa (Sólo no dependientes)
                            pdfFormFields.SetField("Texto41", " ");


                        foreach (var Conyuge in ArrayConyuge)
                        {
                            if (Conyuge == null)
                            {
                                break;
                            }
                            if (Conyuge[1] == Trabajador[6])
                            {
                                switch (Conyuge[2])
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
                                }

                                pdfFormFields.SetField("Texto47", Conyuge[3]); //No. documento Conyuge

                                //Fecha de nacimiento
                                pdfFormFields.SetField("Texto48", Conyuge[9].Substring(6, 4)); // Año
                                pdfFormFields.SetField("Texto49", Conyuge[9].Substring(0, 2)); // Mes 
                                pdfFormFields.SetField("Texto50", Conyuge[9].Substring(3, 2)); // Día

                                //Género Conyuge
                                switch (Conyuge[11])
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
                                switch (Conyuge[26])
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
                                pdfFormFields.SetField("Texto51", Conyuge[15]);

                                //Fecha de Ingreso a la empresa
                                pdfFormFields.SetField("Texto52", " "); // Año
                                pdfFormFields.SetField("Texto53", " "); // Mes 
                                pdfFormFields.SetField("Texto54", " "); // Día

                                pdfFormFields.SetField("Texto57", Conyuge[4] + " " + Conyuge[5]); // Nombres 
                                pdfFormFields.SetField("Texto56", Conyuge[6]); // Primer apellido
                                pdfFormFields.SetField("Texto55", Conyuge[7]); // Segundo apellido

                                pdfFormFields.SetField("Texto58", " "); // Razón social
                                pdfFormFields.SetField("Texto59", " "); // NIT


                                switch ("1") // Recibe subsidio
                                {
                                    case "1": // Si
                                        pdfFormFields.SetField("Casilla de verificación94", "0");
                                        break;

                                    case "2": // No
                                        pdfFormFields.SetField("Casilla de verificación95", "0");
                                        break;

                                    case "NULL":

                                        break;

                                    default:
                                        MessageBox.Show("Error de Recibe subsidio");
                                        break;
                                }

                                pdfFormFields.SetField("Texto60", " "); // Caja por la cual recibe subsidio
                            }

                        }
                        int NumeroBeneficiario = 1;
                        foreach (var Beneficiario in ArrayBeneficiario)
                        {
                            if (Beneficiario == null )
                            {
                                break;
                            }
                            if (Beneficiario[1] == Trabajador[6])
                            {
                                if (NumeroBeneficiario == 1)
                                {
                                    NumeroBeneficiario++;
                                    pdfFormFields.SetField("Texto74", " ");

                                    // 1. Tipo 
                                    //No. Documento de identificación
                                    pdfFormFields.SetField("Texto65", Beneficiario[3]);

                                    //Nombres
                                    pdfFormFields.SetField("Texto34", Beneficiario[5] + " " + Beneficiario[6]);

                                    //Apellidos
                                    pdfFormFields.SetField("Texto61", Beneficiario[7]);

                                    switch (Beneficiario[26]) //Parentesco
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
                                      
                                    }
                                                                 
                                }
                                else if (NumeroBeneficiario == 2)
                                {
                                    NumeroBeneficiario++;
                                    // 2. Tipo 

                                    //No. Documento de identificación
                                    pdfFormFields.SetField("Texto63", Beneficiario[3]);

                                    //Nombres
                                    pdfFormFields.SetField("Texto64", Beneficiario[5] + " " + Beneficiario[6]);

                                    //Apellidos
                                    pdfFormFields.SetField("Texto62", Beneficiario[7]);
                                }
                                else if (NumeroBeneficiario == 3)
                                {
                                    NumeroBeneficiario++;

                                    // 3. Tipo 

                                    //No. Documento de identificación
                                    pdfFormFields.SetField("Texto37", Beneficiario[3]);

                                    //Nombres
                                    pdfFormFields.SetField("Texto66", Beneficiario[5] + " " + Beneficiario[6]);

                                    //Apellidos
                                    pdfFormFields.SetField("Texto69", Beneficiario[7]);
                                }
                                else if (NumeroBeneficiario == 4)
                                {
                                    NumeroBeneficiario++;
                                }
                                else if (NumeroBeneficiario == 5)
                                {
                                    NumeroBeneficiario++;
                                }
                            }
                        }

                        pdfStamper.FormFlattening = true;
                                pdfStamper.Close();
                        
                    }
                          
                    }
                else
                {
                    MessageBox.Show("Error");
                }

            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Seleccione toda la fila");
            //}

        }
    }
}






 
