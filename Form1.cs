using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using FormularioExcel.Modelo;
using iTextSharp.text;
using iTextSharp.text.pdf;



namespace FormularioExcel
{
    public partial class Form1 : Form
    {
        
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
        DataSet errores = new DataSet();
        List<Empresas> Empresas = new List<Empresas>();
        List<Ciudades> ciudadesList = new List<Ciudades>();
        List<Empleador> ArrayEmpleador = new List<Empleador>();
        List<Trabajador> ArrayTrabajador = new List<Trabajador>();
        List<Conyuge> ArrayConyuge = new List<Conyuge>();
        List<Beneficiario> ArrayBeneficiario = new List<Beneficiario>();
        

        public Form1()
        {
            InitializeComponent();

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
           
            adaptador.Fill(ds); 
            conector.Close();

            DataTable a = new DataTable();
            a = ds.Tables[0].Clone();
            errores.Tables.Add(a);

            

            foreach (DataRow Detalles in ds.Tables[0].Rows)
            {
                
                if (Convert.ToString(Detalles[4]) == "0")
                {

                    Empleador Ep = new Empleador();
                    Trabajador tb = new Trabajador();

                    try
                    {
                        Ep.TrabajadorId =  Convert.ToInt32(Detalles[6].ToString());
                        Ep.TipoBeneficiario = Convert.ToInt32(Detalles[4].ToString());
                        Ep.TipoIdentificacionEmpresa = Convert.ToInt32(Detalles[1].ToString());
                        Ep.Nit = Convert.ToInt32(Detalles[2].ToString());
                        Ep.DigitoVerificacion = string.IsNullOrEmpty(Detalles[3].ToString()) ? (int?)null : Convert.ToInt32(Detalles[3].ToString());
                        Ep.FechaRadicacion = Convert.ToDateTime(Detalles[5].ToString());
                        



                        ArrayEmpleador.Add(Ep);

                        tb.trabajadorTipoId = Convert.ToInt32(Detalles[7].ToString());
                        tb.TrabajadorId = Convert.ToInt32(Detalles[6].ToString());
                        tb.TipoBeneficiario = Convert.ToInt32(Detalles[4].ToString());
                        tb.Nombre1 = Detalles[10].ToString(); 
                        tb.Nombre2 = Detalles[11].ToString() == "NULL" ? "" : Detalles[11].ToString();
                        tb.Apellido1 = Detalles[12].ToString();
                        tb.Apellido2 = Detalles[13].ToString() == "NULL"? "" : Detalles[13].ToString();
                        tb.FechaNacimiento = Convert.ToDateTime(Detalles[14].ToString());
                        tb.IdGenero = int.Parse(Detalles[15].ToString());
                        tb.IdNacionalidad = string.IsNullOrEmpty(Detalles[16].ToString()) ? (int?)null : Convert.ToInt32(Detalles[16].ToString()); 
                        tb.IdEstadoCivil = int.Parse(Detalles[17].ToString());
                        tb.IdNivelOcupacion = int.Parse(Detalles[18].ToString()); 
                        tb.FechaIngresoEmpresa = DateTime.Parse(Detalles[19].ToString());
                        tb.HorasMes = int.Parse(Detalles[20].ToString());
                        tb.SalarioBasico = int.Parse(Detalles[21].ToString());
                        tb.Telefono_V= Detalles[22].ToString() == "NULL" ? "" : Detalles[22].ToString();  
                        tb.Celular = Detalles[23].ToString();
                        tb.Direccion_V = Detalles[24].ToString();
                        tb.Dpto_V = int.Parse(Detalles[25].ToString());
                        tb.Ciudad_V = int.Parse(Detalles[26].ToString());
                        tb.Email = Detalles[27].ToString();

                        ArrayTrabajador.Add(tb);
                      
                    }
                    catch (Exception)
                    {
                        errores.Tables[0].ImportRow(Detalles);
                    }

                }
                else if (Convert.ToString(Detalles[4]) == "1")
                {

                    Conyuge cg = new Conyuge();
                    
                    try
                    { 
                        cg.TrabajadorId = Convert.ToInt32(Detalles[6].ToString());
                        cg.TipoBeneficiario = Convert.ToInt32(Detalles[4].ToString());
                        cg.BeneficiarioId = Convert.ToInt32(Detalles[8].ToString());
                        cg.BeneficiarioTipoId = Convert.ToInt32(Detalles[9].ToString());
                        cg.Nombre1 = Detalles[10].ToString();
                        cg.Nombre2 = Detalles[11].ToString() == "NULL" ? "" : Detalles[11].ToString(); 
                        cg.Apellido1 = Detalles[12].ToString();
                        cg.Apellido2 = Detalles[13].ToString() == "NULL" ? "" : Detalles[13].ToString();
                        cg.FechaNacimiento = DateTime.Parse(Detalles[14].ToString());
                        cg.IdGenero = int.Parse(Detalles[15].ToString());
                        cg.FechaIngresoEmpresa = string.IsNullOrEmpty(Detalles[19].ToString()) ? (DateTime?)null : Convert.ToDateTime(Detalles[19].ToString());
                        cg.SalarioBasico = string.IsNullOrEmpty(Detalles[21].ToString()) ? (int?)null : Convert.ToInt32(Detalles[21].ToString());

                        ArrayConyuge.Add(cg);

                    }
                    catch (Exception)
                    {
                        errores.Tables[0].ImportRow(Detalles);

                    }

                }
                else if (Convert.ToString(Detalles[4]) == "2")
                {
                    Beneficiario bn = new Beneficiario();
                    
                    try
                    {
                        bn.TrabajadorId = Convert.ToInt32(Detalles[6].ToString());
                        bn.TipoBeneficiario = Convert.ToInt32(Detalles[4].ToString());
                        bn.BeneficiarioId = Convert.ToInt32(Detalles[8].ToString());
                        bn.BeneficiarioTipoId = Convert.ToInt32(Detalles[9].ToString());
                        bn.Nombre1 = Detalles[10].ToString();
                        bn.Nombre2 = Detalles[11].ToString() == "NULL" ? "" : Detalles[11].ToString(); 
                        bn.Apellido1 = Detalles[12].ToString();
                        bn.Apellido2 = Detalles[13].ToString() == "NULL" ? "" : Detalles[13].ToString();
                        bn.IdParentesco = Convert.ToInt32(Detalles[30].ToString());
                        bn.FechaNacimiento =  Convert.ToDateTime(Detalles[14].ToString()); 
                        bn.IdGenero = int.Parse(Detalles[15].ToString());
                        
                      
                        ArrayBeneficiario.Add(bn);

                    }
                    catch (Exception)
                    {
                        errores.Tables[0].ImportRow(Detalles);

                    }

                }
            }

            
            LblPDF.Text = "Se van a generar " + ArrayTrabajador.Count + " archivos PDF";
            LblPDF.Visible = true;

            LblError.Text = "Se presentaron errores en estas filas:";
            LblError.Visible = true;
            ds = new DataSet();

            
            return errores.Tables[0].DefaultView;
            
        }

        string ruta;

        private void BtnExportar_Click(object sender, EventArgs e)
        {

            //try
            //{   

            OpenFileDialog guardar = new OpenFileDialog();


            if (guardar.ShowDialog() == DialogResult.OK)


                

            //if (MessageBox.Show("Exportar a PDF?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //Lblguardar.Text;

               






                ColsubsidioEntities db = new ColsubsidioEntities();

                    ciudadesList = db.Ciudades.ToList();

                    this.DataDetalles = new DataGridView();
                    ds = new DataSet();
                    int Total = ArrayTrabajador.Count;
                    int Bar = 0;
                
                foreach (var Trabajador in ArrayTrabajador)
                {
                    progressBar.Value = Bar;
                    progressBar.Maximum = Total;
                    progressBar.Increment(Bar);
                    progressBar.Visible = true;
                    Bar ++;
                        
                    string pdfTemplate = System.AppDomain.CurrentDomain.BaseDirectory + "Formulario.pdf";

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
                    string newFile = guardar + Trabajador.Nombre1.ToUpper() + " " + Trabajador.Apellido1.ToUpper() + ".pdf";
                    pdfReader = new PdfReader(pdfTemplate);
                    PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create));
                    AcroFields pdfFormFields = pdfStamper.AcroFields;
                                           
                    Empleador Ep = ArrayEmpleador.Where(f => f.TrabajadorId == Trabajador.TrabajadorId).FirstOrDefault();
                    if (Ep != null)
                    {
                        //1.TIPO DE NOVEDAD

                        //switch ("")
                        //{
                        //    case "1": // Afiliación
                        //        pdfFormFields.SetField("Casilla de verificación1", "0");
                        //        break;

                        //    case "2": // Novedad
                        //        pdfFormFields.SetField("Casilla de verificación2", "0");
                        //        break;

                        //    case "3": // Traslado
                        //        pdfFormFields.SetField("Casilla de verificación3", "0");
                        //        break;

                        //    default:

                        //        break;
                        //}

                        //Fecha de radicación
                        if (Ep.FechaRadicacion != null)
                        {
                            pdfFormFields.SetField("Texto1", Convert.ToString(Ep.FechaRadicacion).Substring(3, 2)); // Día                                                                    
                            pdfFormFields.SetField("Texto2", Convert.ToString(Ep.FechaRadicacion).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto3", Convert.ToString(Ep.FechaRadicacion).Substring(6, 4)); // Año
                        }
                         

                        //3. DATOS DEL EMPLEADOR O ENTIDAD PAGADORA

                        switch (Convert.ToString(Ep.TipoIdentificacionEmpresa)) // Tipo de identidad 
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

                                break;
                        }

                        

                        pdfFormFields.SetField("Texto7", Ep.Nit + "-" + Convert.ToString(Ep.DigitoVerificacion)); // No. de identidad pagadora
                        //pdfFormFields.SetField("Texto9", " "); // Nombre o Razón Social

                        ////Sector
                        //switch ("")
                        //{
                        //    case "1":  // Oficial Público
                        //        pdfFormFields.SetField("Casilla de verificación26", "0");
                        //        break;

                        //    case "2": // Privado
                        //        pdfFormFields.SetField("Casilla de verificación27", "0");
                        //        break;

                        //    case "3": // Mixto
                        //        pdfFormFields.SetField("Casilla de verificación28", "0");
                        //        break;

                        //    default:

                        //        break;
                        //}

                        //pdfFormFields.SetField("Texto10", " "); // Sucursal 
                        //pdfFormFields.SetField("Texto13", " "); // Dirección
                        //pdfFormFields.SetField("Texto12", " "); // Departamento
                        //pdfFormFields.SetField("Texto11", " "); // Teléfono
                    }


                    // 5. DATOS BÁSICOS TRABAJADOR INDEPENDIENTE O PENSIONADO 

                    pdfFormFields.SetField("Texto16", Trabajador.Nombre1.ToUpper() + " " + Trabajador.Nombre2.ToUpper()); // Nombres 
                    pdfFormFields.SetField("Texto17", Trabajador.Apellido1.ToUpper()); // Primer apellido

                   
                        pdfFormFields.SetField("Texto18", Trabajador.Apellido2.ToUpper()); // Segundo apellido
                    
                    

                    switch (Convert.ToString(Trabajador.trabajadorTipoId))
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

                        case "8": // (PT)
                            pdfFormFields.SetField("Casilla de verificación52", "0");
                            break;

                        default:

                            break;
                    }

                    pdfFormFields.SetField("Texto19", Convert.ToString(Trabajador.TrabajadorId)); //No. documento

                    //Estado Civil
                    switch (Convert.ToString(Trabajador.IdEstadoCivil))
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

                            break;
                    }

                    //Fecha de nacimiento
                    pdfFormFields.SetField("Texto20", Convert.ToString(Trabajador.FechaNacimiento).Substring(6, 4)); // Año
                    pdfFormFields.SetField("Texto21", Convert.ToString(Trabajador.FechaNacimiento).Substring(0, 2)); // Mes 
                    pdfFormFields.SetField("Texto22", Convert.ToString(Trabajador.FechaNacimiento).Substring(3, 2)); // Día

                    //Género
                    switch (Convert.ToString(Trabajador.IdGenero))
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

                            break;
                    }

                    //Nivel Ocupacional
                    switch (Convert.ToString(Trabajador.IdNivelOcupacion))
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

                            break;
                    }

                    //Nivel Educativo
                    //switch ("")
                    //{
                    //    case "1": // Ninguno 
                    //        pdfFormFields.SetField("Casilla de verificación61", "0");
                    //        break;

                    //    case "2": // Primaria
                    //        pdfFormFields.SetField("Casilla de verificación62", "0");
                    //        break;

                    //    case "3": // Secundaria
                    //        pdfFormFields.SetField("Casilla de verificación63", "0");
                    //        break;

                    //    case "4": // Técnico
                    //        pdfFormFields.SetField("Casilla de verificación64", "0");
                    //        break;

                    //    case "5": // profesional
                    //        pdfFormFields.SetField("Casilla de verificación65", "0");
                    //        break;

                    //    case "6": // Otro
                    //        pdfFormFields.SetField("Casilla de verificación66", "0");
                    //        break;

                    //    default:

                    //        break;
                    //}

                    //Ingreso a la empresa
                    pdfFormFields.SetField("Texto101", Convert.ToString(Trabajador.FechaIngresoEmpresa).Substring(6, 4)); // Año
                    pdfFormFields.SetField("Texto24", Convert.ToString(Trabajador.FechaIngresoEmpresa).Substring(0, 2)); // Mes 
                    pdfFormFields.SetField("Texto25", Convert.ToString(Trabajador.FechaIngresoEmpresa).Substring(3, 2)); // Día


                    pdfFormFields.SetField("Texto26", Convert.ToString(Trabajador.HorasMes)); // Horas/mes

                    //Trabajador
                    //switch ("")
                    //{
                    //    case "1": // urbano(UR) 
                    //        pdfFormFields.SetField("Casilla de verificación74", "0");
                    //        break;

                    //    case "2": // Rural(Ru)
                    //        pdfFormFields.SetField("Casilla de verificación75", "0");
                    //        break;

                    //    default:

                    //        break;
                    //}

                    //Salario básico/mesada
                    pdfFormFields.SetField("Texto27", Convert.ToString(Trabajador.SalarioBasico));

                    //Celular
                    pdfFormFields.SetField("Texto28", Trabajador.Celular);

                    ////EPS (Solo para independientes)
                    //pdfFormFields.SetField("Texto29", " ");

                    ////AFP (Solo para independientes)
                    //pdfFormFields.SetField("Texto30", " ");

                    //Dirección vivienda
                    pdfFormFields.SetField("Texto100", Trabajador.Direccion_V.ToUpper());

                    
                    ////Departamento
                    //switch (Trabajador[25])
                    //{
                    //    case "11":
                    //        pdfFormFields.SetField("Texto32", "BOGOTA");
                    //        break;

                    //    default:

                    //        break;
                    //}

                    //Zona vivienda
                    //switch ("")
                    //{
                    //    case "1": // urbano(UR) 
                    //        pdfFormFields.SetField("Casilla de verificación77", "0");
                    //        break;

                    //    case "2": // Rural(Ru)
                    //        pdfFormFields.SetField("Casilla de verificación76", "0");
                    //        break;

                    //    default:

                    //        break;
                    //}

                    //Teléfono vivienda
                    pdfFormFields.SetField("Texto31", Trabajador.Telefono_V);

                    ////Dirección trabajo
                    //pdfFormFields.SetField("Texto38", " ");

                    ////Municipio
                    //pdfFormFields.SetField("Texto36", " ");

                    ////Departamento
                    //pdfFormFields.SetField("Texto102", " ");

                    //Zona Trabajo
                    //switch ("")
                    //{
                    //    case "1": // urbano(UR) 
                    //        pdfFormFields.SetField("Casilla de verificación78", "0");
                    //        break;

                    //    case "2": // Rural(Ru)
                    //        pdfFormFields.SetField("Casilla de verificación79", "0");
                    //        break;

                    //    default:

                    //        break;
                    //}

                    //Teléfono trabajo 
                    //pdfFormFields.SetField("Texto33", " ");

                    //Correo Electrónico
                    pdfFormFields.SetField("Texto39", Trabajador.Email.ToUpper());

                    //País y Ciudad de Residencia (Solo para Residentes en el Exterior)
                    //pdfFormFields.SetField("Texto40", " ");

                    ////Programa (Sólo no dependientes)
                    //pdfFormFields.SetField("Texto41", " ");


                    Ciudades ciudad = ciudadesList.Where(f => f.Ciudad_Id == Trabajador.Ciudad_V.Value).FirstOrDefault();
                    Ciudades Dpto = ciudadesList.Where(f => f.CodDepartamento == Trabajador.Dpto_V.Value).FirstOrDefault();
                    if (ciudad != null)
                    {
                        pdfFormFields.SetField("Texto35", ciudad.Nombre); //Ciudad Trabajador
                        
                    }
                    if (Dpto != null)
                    {
                        pdfFormFields.SetField("Texto32", ciudad.Departamento); //Departamento trabajador
                    }

                    Conyuge cg= ArrayConyuge.Where(f => f.TrabajadorId == Trabajador.TrabajadorId).FirstOrDefault();

                    //foreach (var Conyuge in A)
                    //{
                        if (cg != null)
                        {

                            //7.Información del cónyuge o compañer@
                            switch (Convert.ToString(cg.BeneficiarioTipoId))
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

                                    break;
                            }

                            pdfFormFields.SetField("Texto47", Convert.ToString(cg.BeneficiarioId)); //No. documento Conyuge

                            //Fecha de nacimiento
                            
                            
                            pdfFormFields.SetField("Texto48", Convert.ToString(cg.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto49", Convert.ToString(cg.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto50", Convert.ToString(cg.FechaNacimiento).Substring(3, 2)); // Día
                              
                            

                            //Género Conyuge
                            switch (Convert.ToString(cg.IdGenero))
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

                                    break;
                            }

                            //Trabaja Conyuge
                            switch (cg.TrabajaConyugue)
                            {
                                case "SI": // Si
                                    pdfFormFields.SetField("Casilla de verificación92", "0");
                                    break;

                                case "NO": // No
                                    pdfFormFields.SetField("Casilla de verificación93", "0");
                                    break;

                                case "NULL":
                                    pdfFormFields.SetField("Casilla de verificación92", "");
                                    pdfFormFields.SetField("Casilla de verificación93", "");
                                    break;

                                default:

                                    break;
                            }

                            //salario básico conyuge 
                            pdfFormFields.SetField("Texto51", Convert.ToString(cg.SalarioBasico));

                        //Fecha de Ingreso a la empresa
                            if (cg.FechaIngresoEmpresa != null)
                            { 
                                pdfFormFields.SetField("Texto52", Convert.ToString(cg.FechaIngresoEmpresa).Substring(6, 4)); // Año
                                pdfFormFields.SetField("Texto53", Convert.ToString(cg.FechaIngresoEmpresa).Substring(0, 2)); // Mes 
                                pdfFormFields.SetField("Texto54", Convert.ToString(cg.FechaIngresoEmpresa).Substring(3, 2)); // Día
                            }
                        

                            pdfFormFields.SetField("Texto57", cg.Nombre1.ToUpper() + " " + cg.Nombre2.ToUpper()); // Nombres 
                            pdfFormFields.SetField("Texto56", cg.Apellido1.ToUpper()); // Primer apellido
                            pdfFormFields.SetField("Texto55", cg.Apellido2.ToUpper()); // Segundo apellido

                            //pdfFormFields.SetField("Texto58", " "); // Razón social
                            //pdfFormFields.SetField("Texto59", " "); // NIT


                            //switch ("") // Recibe subsidio
                            //{
                            //    case "1": // Si
                            //        pdfFormFields.SetField("Casilla de verificación94", "0");
                            //        break;

                            //    case "2": // No
                            //        pdfFormFields.SetField("Casilla de verificación95", "0");
                            //        break;

                            //    case "NULL":

                            //        break;

                            //    default:

                            //        break;
                            //}

                            //pdfFormFields.SetField("Texto60", " "); // Caja por la cual recibe subsidio

                        }

                    //}

                    //8. INFORMACIÓN GRUPO FAMILIAR 

                    List<Beneficiario> bf = ArrayBeneficiario.Where(f => f.TrabajadorId == Trabajador.TrabajadorId).ToList();

                    int numerob = 1;
                    foreach (var beneficiarios in bf)
                    {
                        if (numerob == 1)
                        {
                            // 1. Tipo 
                            switch (Convert.ToString(beneficiarios.BeneficiarioTipoId))
                            {
                                case "1": //Cédula de ciudadania
                                    pdfFormFields.SetField("Texto74", "CC");
                                    break;

                                case "2": //Carnet
                                    pdfFormFields.SetField("Texto74", "CD");
                                    break;

                                case "3": //Cedula extranjera
                                    pdfFormFields.SetField("Texto74", "CE");
                                    break;

                                case "4": //Pasaporte 
                                    pdfFormFields.SetField("Texto74", "PA");
                                    break;

                                case "5": //Tarjeta de identidad  
                                    pdfFormFields.SetField("Texto74", "TI");
                                    break;

                                case "6": //Permiso especial
                                    pdfFormFields.SetField("Texto74", "PE");
                                    break;

                                case "8": //Registro
                                    pdfFormFields.SetField("Texto74", "RC");
                                    break;

                                case "9": //Permiso de turismo
                                    pdfFormFields.SetField("Texto74", "PT");
                                    break;

                                default:

                                    break;
                            }

                            //No. Documento de identificación
                            pdfFormFields.SetField("Texto65", Convert.ToString(beneficiarios.BeneficiarioId));

                            //Nombres
                            pdfFormFields.SetField("Texto34", beneficiarios.Nombre1.ToUpper() + " " + beneficiarios.Nombre2.ToUpper());

                            //Apellidos
                            pdfFormFields.SetField("Texto61", beneficiarios.Apellido1.ToUpper() + " " + beneficiarios.Apellido2.ToUpper());

                            switch (Convert.ToString(beneficiarios.IdParentesco)) //Parentesco
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

                                    break;
                            }

                            //Fecha de Nacimiento 
                            pdfFormFields.SetField("Texto91", Convert.ToString(beneficiarios.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto93", Convert.ToString(beneficiarios.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto92", Convert.ToString(beneficiarios.FechaNacimiento).Substring(3, 2)); // Día

                            numerob++;
                        }
                        else if (numerob == 2)
                        {
                          //2.Tipo
                            switch (Convert.ToString(beneficiarios.BeneficiarioTipoId))
                            {
                                case "1": //Cédula de ciudadania
                                    pdfFormFields.SetField("Texto75", "CC");
                                    break;

                                case "2": //Carnet
                                    pdfFormFields.SetField("Texto75", "CD");
                                    break;

                                case "3": //Cedula extranjera
                                    pdfFormFields.SetField("Texto75", "CE");
                                    break;

                                case "4": //Pasaporte 
                                    pdfFormFields.SetField("Texto75", "PA");
                                    break;

                                case "5": //Tarjeta de identidad  
                                    pdfFormFields.SetField("Texto75", "TI");
                                    break;

                                case "6": //Permiso especial
                                    pdfFormFields.SetField("Texto75", "PE");
                                    break;

                                case "8": //Registro
                                    pdfFormFields.SetField("Texto75", "RC");
                                    break;

                                case "9": //Permiso de turismo
                                    pdfFormFields.SetField("Texto75", "PT");
                                    break;

                                default:

                                    break;
                            }

                            //No. Documento de identificación
                            pdfFormFields.SetField("Texto63", Convert.ToString(beneficiarios.BeneficiarioId));

                            //Nombres
                            pdfFormFields.SetField("Texto64", beneficiarios.Nombre1.ToUpper() + " " + beneficiarios.Nombre2.ToUpper());

                            //Apellidos
                            pdfFormFields.SetField("Texto62", beneficiarios.Apellido1.ToUpper() + " " + beneficiarios.Apellido2.ToUpper());

                            switch (Convert.ToString(beneficiarios.IdParentesco)) //Parentesco
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

                                    break;
                            }

                            //Fecha de Nacimiento 
                            pdfFormFields.SetField("Texto88", Convert.ToString(beneficiarios.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto90", Convert.ToString(beneficiarios.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto89", Convert.ToString(beneficiarios.FechaNacimiento).Substring(3, 2)); // Día

                            numerob++;
                        }
                        else if (numerob == 3)
                        {
                            // 3. Tipo 
                            switch (Convert.ToString(beneficiarios.BeneficiarioTipoId))
                            {
                                case "1": //Cédula de ciudadania
                                    pdfFormFields.SetField("Texto76", "CC");
                                    break;

                                case "2": //Carnet
                                    pdfFormFields.SetField("Texto76", "CD");
                                    break;

                                case "3": //Cedula extranjera
                                    pdfFormFields.SetField("Texto76", "CE");
                                    break;

                                case "4": //Pasaporte 
                                    pdfFormFields.SetField("Texto76", "PA");
                                    break;

                                case "5": //Tarjeta de identidad  
                                    pdfFormFields.SetField("Texto76", "TI");
                                    break;

                                case "6": //Permiso especial
                                    pdfFormFields.SetField("Texto76", "PE");
                                    break;

                                case "8": //Registro
                                    pdfFormFields.SetField("Texto76", "RC");
                                    break;

                                case "9": //Permiso de turismo
                                    pdfFormFields.SetField("Texto76", "PT");
                                    break;

                                default:

                                    break;
                            }

                            //No. Documento de identificación
                            pdfFormFields.SetField("Texto37", Convert.ToString(beneficiarios.BeneficiarioId));

                            //Nombres
                            pdfFormFields.SetField("Texto66", beneficiarios.Nombre1.ToUpper() + " " + beneficiarios.Nombre2.ToUpper());

                            //Apellidos
                            pdfFormFields.SetField("Texto69", beneficiarios.Apellido1.ToUpper() + " " + beneficiarios.Apellido2.ToUpper());

                            switch (Convert.ToString(beneficiarios.IdParentesco)) //Parentesco
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

                                    break;
                            }

                            //Fecha de Nacimiento 
                            pdfFormFields.SetField("Texto85", Convert.ToString(beneficiarios.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto87", Convert.ToString(beneficiarios.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto86", Convert.ToString(beneficiarios.FechaNacimiento).Substring(3, 2)); // Día

                            numerob++;
                        }
                        else if (numerob == 4)
                        {
                            //4.Tipo
                            switch (Convert.ToString(beneficiarios.BeneficiarioTipoId))
                            {
                                case "1": //Cédula de ciudadania
                                    pdfFormFields.SetField("Texto77", "CC");
                                    break;

                                case "2": //Carnet                
                                    pdfFormFields.SetField("Texto77", "CD");
                                    break;

                                case "3": //Cedula extranjera     
                                    pdfFormFields.SetField("Texto77", "CE");
                                    break;

                                case "4": //Pasaporte             
                                    pdfFormFields.SetField("Texto77", "PA");
                                    break;

                                case "5": //Tarjeta de identidad  
                                    pdfFormFields.SetField("Texto77", "TI");
                                    break;

                                case "6": //Permiso especial
                                    pdfFormFields.SetField("Texto77", "PE");
                                    break;

                                case "8": //Registro
                                    pdfFormFields.SetField("Texto77", "RC");
                                    break;

                                case "9": //Permiso de turismo
                                    pdfFormFields.SetField("Texto77", "PT");
                                    break;

                                default:

                                    break;
                            }

                            //No. Documento de identificación
                            pdfFormFields.SetField("Texto67", Convert.ToString(beneficiarios.BeneficiarioId));

                            //Nombres
                            pdfFormFields.SetField("Texto70", beneficiarios.Nombre1.ToUpper() + " " + beneficiarios.Nombre2.ToUpper());

                            //Apellidos
                            pdfFormFields.SetField("Texto68", beneficiarios.Apellido1.ToUpper() + " " + beneficiarios.Apellido2.ToUpper());

                            switch (Convert.ToString(beneficiarios.IdParentesco)) //Parentesco
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

                                    break;
                            }

                            //Fecha de Nacimiento 
                            pdfFormFields.SetField("Texto82", Convert.ToString(beneficiarios.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto84", Convert.ToString(beneficiarios.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto83", Convert.ToString(beneficiarios.FechaNacimiento).Substring(3, 2)); // Día

                            numerob++;
                        }
                        else
                        {
                            // 5. Tipo 
                            switch (Convert.ToString(beneficiarios.BeneficiarioTipoId))
                            {
                                case "1": //Cédula de ciudadania
                                    pdfFormFields.SetField("Texto78", "CC");
                                    break;

                                case "2": //Carnet
                                    pdfFormFields.SetField("Texto78", "CD");
                                    break;

                                case "3": //Cedula extranjera
                                    pdfFormFields.SetField("Texto78", "CE");
                                    break;

                                case "4": //Pasaporte 
                                    pdfFormFields.SetField("Texto78", "PA");
                                    break;

                                case "5": //Tarjeta de identidad  
                                    pdfFormFields.SetField("Texto78", "TI");
                                    break;

                                case "6": //Permiso especial
                                    pdfFormFields.SetField("Texto78", "PE");
                                    break;

                                case "8": //Registro
                                    pdfFormFields.SetField("Texto78", "RC");
                                    break;

                                case "9": //Permiso de turismo
                                    pdfFormFields.SetField("Texto78", "PT");
                                    break;

                                default:

                                    break;
                            }

                            //No. Documento de identificación
                            pdfFormFields.SetField("Texto71", Convert.ToString(beneficiarios.BeneficiarioId));

                            //Nombres
                            pdfFormFields.SetField("Texto72", beneficiarios.Nombre1.ToUpper() + " " + beneficiarios.Nombre2.ToUpper());

                            //Apellidos
                            pdfFormFields.SetField("Texto73", beneficiarios.Apellido1.ToUpper() + " " + beneficiarios.Apellido2.ToUpper());

                            switch (Convert.ToString(beneficiarios.IdParentesco)) //Parentesco
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

                                    break;
                            }

                            //Fecha de Nacimiento 
                            pdfFormFields.SetField("Texto79", Convert.ToString(beneficiarios.FechaNacimiento).Substring(6, 4)); // Año
                            pdfFormFields.SetField("Texto80", Convert.ToString(beneficiarios.FechaNacimiento).Substring(0, 2)); // Mes 
                            pdfFormFields.SetField("Texto81", Convert.ToString(beneficiarios.FechaNacimiento).Substring(3, 2)); // Día

                            numerob++;
                        }

                         
                    }
                        //AUTORIZACIÓN PARA LA UTILIZACIÓN DE DATOS PERSONALES
                         pdfFormFields.SetField("Casilla de verificación121", "0");

                         pdfStamper.FormFlattening = true;
                         pdfStamper.Close();
                }
                    //db.Empleador.AddRange(ArrayEmpleador);
                    //db.Trabajador.AddRange(ArrayTrabajador);
                    //db.SaveChanges();

                    //db.Beneficiario.AddRange(ArrayBeneficiario);
                    //db.Conyuge.AddRange(ArrayConyuge);
                    //db.SaveChanges();



                MessageBox.Show("Finalizo, se generaron " + Bar + " formatos de colsubsidio");
            }

            else
            {
                MessageBox.Show("Error");
            }

            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Algo salio mal verifique");
            //}

}
    }
}

