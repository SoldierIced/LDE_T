using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml.Linq;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using System.Collections;
using System.Data.OleDb;


namespace LectorXML
{
    public partial class w : Form
    { 

        
     
        //MessageBox.Show(documento.ToString());
        Materias []M;
        PDFcreater PDF= new PDFcreater();
        int vuelta = 0;
        
        string destino_archivo = "";


        XDocument documento;
        public w()
        {
           

            InitializeComponent();
           

        }

        public int NivelRespuesta(string respuesta)
        {


            switch (respuesta)
            {

                case "Excelente":
                    return 1;
                case "No faltó nunca":
                    return 1;
                case "Siempre puntual":
                    return 1;
                case "Siempre cumple":
                    return 1;
                case "Siempre":
                    return 1;
                case "Excelente relación":
                    return 1;
                case "Inmediata":
                    return 1;
                case "Siempre responde":
                    return 1;


                case "Muy bueno":
                    return 2;
                case "A veces faltó":
                    return 2;

                case "A veces llegó tarde" :
            return 2;

                case "A veces termina antes":
                    return 2;

                case "Casi siempre":
                    return 2;

                case "Casi siempre responde":
                    return 2;

                case "Algunos temas no":
                    return 2;

                case "Muy rápida":
                    return 2;

                case "Bueno":
                    return 3;

                case "Faltó varias veces":
                    return 3;

                case "Llegó tarde muchas veces":
                    return 3;

                case "Muchas veces termina antes":
                    return 3;

                case "Pocas veces":
                    return 3;

                case "Varios temas no":
                    return 3;

                case "Rápida":
                    return 3;

                case "Pocas veces responde":
                    return 3;

                case "Regular":
                    return 4;

                case "Faltó muchas veces":
                    return 4;

                case "Casi siempre llegó tarde":
                    return 4;

                case "Casi siempre termina antes":
                    return 4;

                case "Casi nunca":
                    return 4;

                case "Muchos temas no":
                    return 4;

                case "Lenta":
                    return 4;

                case "Casi nunca responde":
                    return 4;

                case "Malo":
                    return 5;

                case "Casi nunca vino":
                    return 5;

                case "Nunca llegó a horario":
                    return 5;

                case "Siempre se retira antes":
                    return 5;

                case "Siempre termina antes":
                    return 5;

                case "Nunca":
                    return 5;

                case "Ninguna relación":
                    return 5;

                case "Muy lenta":
                    return 5;

                case "Nunca responde":
                    return 5;


                default:

                    return 0;



            }

                    }




        private void Form1_Load(object sender, EventArgs e)
        {
            //    Login lg = new Login();

            //  lg.ShowDialog();
           
        }
        public void cargarlistbox()
        {
           
            var usuarios = from usu in documento.Descendants("VFPData") select usu;
            foreach (XElement u in usuarios.Elements("_exportar"))
            {
              //  MessageBox.Show((u.Element("respuesta").Value));
                if (Revisarlistbox(u.Element("año").Value, listBox10) == -1)
                    listBox10.Items.Add(u.Element("año").Value);

                if (Revisarlistbox(u.Element("legajo").Value, listBox9) == -1)
                    listBox9.Items.Add(u.Element("legajo").Value);

                if (Revisarlistbox(u.Element("esp").Value, listBox8) == -1)
                    listBox8.Items.Add(u.Element("esp").Value);

                if (Revisarlistbox(u.Element("materia").Value, listBox7) == -1)
                    listBox7.Items.Add(u.Element("materia").Value);

                if (Revisarlistbox(u.Element("plan").Value, listBox6) == -1)
                    listBox6.Items.Add(u.Element("plan").Value);

                if (Revisarlistbox(u.Element("comision").Value, listBox5) == -1)
                    listBox5.Items.Add(u.Element("comision").Value);

                if (Revisarlistbox(u.Element("curso").Value, listBox4) == -1)
                    listBox4.Items.Add(u.Element("curso").Value);

                if (Revisarlistbox(u.Element("encuesta").Value, listBox3) == -1)
                    listBox3.Items.Add(u.Element("encuesta").Value);

                if (Revisarlistbox(u.Element("numero_de_").Value, listBox2) == -1)
                    listBox2.Items.Add(u.Element("numero_de_").Value);

                if (Revisarlistbox(u.Element("preguntas").Value, listBox1) == -1)
                    listBox1.Items.Add(u.Element("preguntas").Value);



               try { 
                if (Revisarlistbox(u.Element("respuesta").Value, listBox11) == -1)
                    listBox11.Items.Add(u.Element("respuesta").Value);



                }
                catch (Exception aq) {
                   
                    
                }
            }
        }

        public int Revisarlistbox(string valor, ListBox asd)
        {

            for (int i = 0; i < asd.Items.Count; i++)
            {
                // MessageBox.Show("valor es > " + valor + "comparando con ; " + listBox1.Items[i].ToString());
                if (valor == asd.Items[i].ToString()) { return i; }
                //MessageBox.Show(listBox1.Items[i].ToString());
            }



            return -1;
        }
        public int Verificarmateria(XElement u)
        {
         
            string a = "";
            try { 
            a =
                      u.Element("materia").Value + "<>" +

                      u.Element("comision").Value + "<>" +
                      u.Element("curso").Value;
            }
            catch (Exception qa)
            {
                a =
                      u.Element("materia").Value + "<>" +

                      u.Element("comision").Value + "<>" +
                      u.Element("curso").Value;

            }

            return Revisarlistbox(a, listBox20);
              

                
        }
        public void cargarmaterias() 
        {
            var usuarios = from usu in documento.Descendants("VFPData") select usu;
            string a = "";
            foreach (XElement u in usuarios.Elements("_exportar"))
            {
               
                try {
                    a =
                             u.Element("materia").Value + "<>" +

                             u.Element("comision").Value + "<>" +
                             u.Element("curso").Value;

                }
                catch (Exception aq)
                {
                    a =
                         u.Element("materia").Value + "<>" +

                         u.Element("comision").Value + "<>" +
                         u.Element("curso").Value;
                }
                if (Revisarlistbox(a, listBox20) == -1) { 


                listBox20.Items.Add(a);
                }

            }


            }
      


            public void contadorM()
        {

          //  MessageBox.Show(listBox2.Items.Count.ToString() + listBox11.Items.Count.ToString());
            
            int i, a, x;
            for (i = 0; i < listBox20.Items.Count; i++) //materias
            {   
                
                M[i] = new Materias(listBox2.Items.Count, listBox11.Items.Count); //respuesta
                
            }
            // MessageBox.Show("valor de i > "+i.ToString());
            var usuarios = from usu in documento.Descendants("VFPData") select usu;
            a = 0;
            x = 0;
            foreach (XElement u in usuarios.Elements("_exportar"))
            {
                a = Verificarmateria(u);
                // MessageBox.Show("valor de a > "+a.ToString());
                if (a != -1)
                   // MessageBox.Show(a.ToString());
                {
                  
                    try {
                      //  aa = int.Parse(u.Element("materia").Value) ;
                        x = int.Parse(u.Element("numero_de_").Value) - 1;
                       // MessageBox.Show(a+"      "+x);
                        M[a].anio = (u.Element("año").Value);
                        M[a].esp = (u.Element("esp").Value);
                        M[a].curso = (u.Element("curso").Value);
                        M[a].COD = (u.Element("materia").Value);//
                    }
                    catch (Exception ex) {
                        MessageBox.Show(ex.ToString());
                    }
                    

                    try {
                        if (u.Element("observacio").Value != "") {
                            if (u.Element("respuesta").Value != "") {
                               
                                M[a].CargarObser(u.Element("observacio").Value, x, u.Element("respuesta").Value, (u.Element("preguntas").Value));
                               
                            }
                         

                        }
                     


                    }

                    catch(Exception sq) { }


                    //   MessageBox.Show("valor de funcion parse" + (int.Parse(u.Element("numero_de_").Value) - 1).ToString());
                    try
                    {
                        M[a].m[x, Revisarlistbox(u.Element("respuesta").Value, listBox11)] += 1;
                    }
                    catch (Exception aq) {
                    
                    }

                }


                //[cantidadPreguntas, cantidadrespuestasP];


            }


        }

        public int contarRespuestas(string cod, string esp, string anio,string curso)
        {
            var usuarios = from usu in documento.Descendants("VFPData") select usu;
            int cont = 0;
            
            foreach (XElement u in usuarios.Elements("_exportar"))
            {

                try {

                    if (u.Element("materia").Value == cod)
                    {
                        if (u.Element("año").Value == anio)
                        {
                            if (u.Element("esp").Value == esp)
                            {
                                if (u.Element("curso").Value == curso)
                                {
                                    //falta esto.
                                    //
                                   if (Revisarlistbox(u.Element("legajo").Value, listBox12) == -1) {

                                        listBox12.Items.Add(u.Element("legajo").Value);
                                        cont++;
                                    }

                                }

                            }
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                  
                
            }
            return cont;
        }


        public float porcentaje (int valor, int total)
        {
            float porce = 0;

            porce = (valor * 100) / total;
            porce = float.Parse( Math.Round(porce,1).ToString());
            
            return porce;
        }




        public int cargarcolores (int aq)
        {

           int  poin = 0;
            chart1.Series[0].Points.Clear();


            for (int i = 0; i < listBox11.Items.Count; i++) //respuestas
            {                                if (listBox11.Items[i].ToString() != "")
                {
                    chart1.Titles.Clear();
                    chart1.Titles.Add(listBox1.Items[vuelta].ToString());  //preguntas
                                                                          
                    if (M[aq].m[vuelta, i] != 0)
                    {
                        poin++;
                    }
                }
                else
                {                    M[aq].m[vuelta, i] = 0;
                }
            }

            return poin;


        }





      public void CompletarPDF ()
        {
            int poin = 0;
            int porce=10;
           
           
            

            for (int aq = 0; aq < listBox20.Items.Count; aq++)
            {
                progressBar1.Increment(100);

                //  MessageBox.Show((listBox7.Items.Count/10 ==aq).ToString());
                if (listBox20.Items.Count / porce == aq)
                {
                    
                    porce += 10;
                    progressBar1.Increment(7);
                   
                }
                 try
                {
                    PDF.CadenaGuardar = @txt_direccarpe.Text;
                    string name = "";
                    string namecar = "";
                    BuscarNombreMateria(M[aq].COD, M[aq].esp, ref name, ref namecar);
                    M[aq].respuestas = contarRespuestas(M[aq].COD, M[aq].esp, M[aq].anio, M[aq].curso);
                    PDF.CargaP("Materia Nro." + M[aq].COD,M[aq],name,namecar,M[aq].respuestas);

                    for (vuelta = 0; vuelta < listBox1.Items.Count; vuelta++) //preguntas
                    {
                        int[,] Resp = new int[5, 5];
                        
                           
                        poin = 0;

                        chart1.Series["Series1"].Points.Clear();

                       chart1.Series["Series1"]["PieLabelStyle"] = "Outside";
                        

                        for (int i = 0; i < listBox11.Items.Count; i++) //respuestas
                        {
                             
                            if (listBox11.Items[i].ToString()!= "") {
                                try {    
                                chart1.Titles.Clear();
                                chart1.Titles.Add(listBox1.Items[vuelta].ToString());  //preguntas
                                    System.Drawing.Font aw = new System.Drawing.Font("Microsoft Sans Serif",12);

                                    chart1.Titles[0].Font = aw;
                                        //valor numerico a tomar para grafico.
                                    if (M[aq].m[vuelta, i] != 0)
                                {
                                    if ((aq * 100) / (listBox20.Items.Count - 1) == 10)
                                    {
                                      

                                    }

                                    
                                    switch (NivelRespuesta(listBox11.Items[i].ToString()))
                                    {
                                        case 1:
                                            poin = 0;
                                           
                                            break;
                                        case 2:
                                            poin = 1;
                                            
                                            break;
                                        case 3:
                                            poin = 2;
                                           
                                            break;
                                        case 4:
                                            poin = 3;
                                           
                                            break;
                                        case 5:
                                            poin = 4;
                                           
                                            break;
                                        default:
                                            break;


                                    }


                                        Resp[poin, 0] = 1;
                                    Resp[poin, 2] = i;
                                    Resp[poin, 3] = M[aq].m[vuelta, i];
                                       
                                }

                                }
                                catch (Exception excep) {
                                    MessageBox.Show("error arriba ");
                                    MessageBox.Show(excep.ToString());

                                }


                            }
                            else
                            {
                                M[aq].m[vuelta, i] = 0;
                            }


                                

                        }



                        //System.Drawing.Font aw = new System.Drawing.Font("Microsoft Sans Serif", 8);
                        int cont = 0; ;
                        for (int sa =0; sa <5; sa++)
                        {
                            // if (Resp[sa, 0]==1) { 

                            // MessageBox.Show(Resp[sa, 0].ToString() + Resp[sa, 2] + Resp[sa, 3]);
                                                       if (Resp[sa,3] !=0) {
                                
                                try
                                {
                                   // chart1.Series["Series1"]["LabelStyle"] = 
                                        chart1.Series["Series1"].Points.Add();
                                 /*   chart1.TextAntiAliasingQuality = TextAntiAliasingQuality.High;
                                    chart1.AntiAliasing = AntiAliasingStyles.All;
                                    chart1.IsSoftShadows = true;*/
                                        chart1.Series["Series1"].Points[cont].SetValueY(Resp[sa,3]);
                                    chart1.Series["Series1"].Points[cont].CustomProperties = "LabelsHorizontalLineSize=1.5, PieLabelStyle=Outside, LabelsRadialLineSize=0, PieLineColor=Black";

                                    chart1.Series["Series1"].Points[cont].AxisLabel = Resp[sa, 3].ToString();// porcentaje(M[aq].m,Resp[sa,3], vuelta, listBox11).ToString() + " %";
                                                                                                             // chartMarcas.Series[0]["PieLabelStyle"] = "Outside";

                                    //  chart1.Series["Series1"].Points[cont].AxisLabel.
                                    //  chart1.Series["Series1"].Points[cont].Font = aw;

                                    // nombre del conjunto (serian los tipos de respuestas)

                                    chart1.Series["Series1"].Points[cont].LegendText = listBox11.Items[Resp[sa, 2]].ToString();// +" ("+Resp[sa, 3]+")";

                                    switch (sa+1)
                                    {
                                        case 1:
                                          
                                            chart1.Series["Series1"].Points[cont].Color = Color.FromArgb(33,89,103);

                                            break;
                                        case 2:
                                           
                                            chart1.Series["Series1"].Points[cont].Color = Color.FromArgb(0,176,80);

                                            break;
                                        case 3:
                                            chart1.Series["Series1"].Points[cont].Color = Color.FromArgb(148,208,80);

                                            break;
                                        case 4:
                                            chart1.Series["Series1"].Points[cont].Color = Color.FromArgb(255, 255, 0);

                                            break;
                                        case 5:
                                            chart1.Series["Series1"].Points[cont].Color = Color.FromArgb(230,0,0);

                                            break;
                                        default:
                                          
                                            break;


                                    }
                                }
                                catch (Exception excep)
                                {
                                    MessageBox.Show(excep.ToString());
                                    MessageBox.Show(sa.ToString() + "  Vuelta:  " + vuelta + "  Respo 2: " + Resp[sa, 2] + "Resp 3:  " + Resp[sa, 3]);



                                }
                                    cont++;
                            }
                                                        }
                      //  MessageBox.Show(cont.ToString());
                        int total=0;
                        for (int q = 0; q < cont; q++) {
                           // MessageBox.Show(chart1.Series["Series1"].Points[q].AxisLabel);
                            total += int.Parse(chart1.Series["Series1"].Points[q].AxisLabel);

                        }

                       

                        try {
                            decimal ar;
                            for (int q = 0; q < cont; q++)
                            {
                                
                                ar = int.Parse(chart1.Series["Series1"].Points[q].AxisLabel) * 100;
                                ar = ar / total;
                                ar = Math.Round(ar,0);
                                chart1.Series["Series1"].Points[q].AxisLabel = ar.ToString()+ "%";


                            }
                        }
                        catch (Exception aw) { MessageBox.Show(aw.ToString()); }


                        if (PDF.cont != 1) { PDF.x += 300; }
                        if ((PDF.cont) == 3) { PDF.y -=180; PDF.x =10; PDF.cont = 1; }

                        PDF.Guardar(chart1);

                        PDF.cont++;
                      

                          
                    }
                   
                   
                    for (int jj=0; jj <M[aq].cantpreg; jj++) {
                       
                        
                        if (M[aq].obs[jj]!= "") {
 
                            PDF.agregarObser(M[aq].preg[jj],1);
                            PDF.agregarObser("\n" + M[aq].obs[jj],0);

                        }
                    }
                    PDF.CerrarPDF();
                    PDF.conta = 0;
                   
                }

                catch (Exception aw)
                {
                    MessageBox.Show("Error en Lectura de archivo, Puede ser que alguien lo este usando. Porfavor, cierrelo.  "+ "\n"+
                        "la encuesta de esta materia no se ha generado.");
                }

               
                
            }

            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }

    private void chart1_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog(); 

            openFileDialog1.InitialDirectory = "G:\\PROYECTO GERARDO\\Encuestas";
            
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // MessageBox.Show(openFileDialog1.FileName.ToString());// Insert code to read the stream here.

                             destino_archivo = openFileDialog1.FileName.ToString();
                            txt_direxml.Text = destino_archivo;
                            txt_direccarpe.Text = Path.GetDirectoryName(openFileDialog1.FileName) + @"\";
                            //MessageBox.Show(destino);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
               //string carpeta =txt_direccarpe.Text+ "Encuesta, dia ";
                //MessageBox.Show(carpeta);
            }
        }

        private void btn_cargaDatos_Click(object sender, EventArgs e)
        {
            
            try
            {
                label3.Visible = true;

                if (txt_direccarpe.Text!= string.Empty) {


                    documento = XDocument.Load(@txt_direxml.Text);
                  


                    cargarlistbox();

                    listBox11.Items.Add("");
                    cargarmaterias();


                    //  MessageBox.Show(progressBar1.Value.ToString());
                    timer1.Start();
                    progressBar1.Value = 0;
                    progressBar1.Maximum = (100 * listBox20.Items.Count)+15;
                    label3.Visible = true;

                    progressBar1.Visible = true;
                    progressBar1.Value += 15;





                    //  MessageBox.Show(listBox7.Items.Count.ToString());
                    M = new Materias[listBox20.Items.Count];

                
                    contadorM();
                    progressBar1.Increment(20);

                    //  MessageBox.Show(progressBar1.Value.ToString());

                    PDF.CadenaGuardar = @txt_direccarpe.Text;


                    chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
                    CompletarPDF();
                    progressBar1.Value = progressBar1.Maximum;
                  

                 
                    //  MessageBox.Show(progressBar1.Value.ToString());

                }
                else
                {


                    MessageBox.Show("Recuerde poner direccion de en donde desea guardar los PDFs.");

                }

                borrartodo();
                label3.Text = "Pdf Ya creados.";
            }
            catch (Exception aw)
            {

                MessageBox.Show(aw.ToString());
                MessageBox.Show("Verifique si puso bien las direcciones.");


            }

           
        }

        public void borrartodo()
        {

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();
            listBox7.Items.Clear();
            listBox8.Items.Clear();
            listBox9.Items.Clear();
            listBox10.Items.Clear();
            listBox11.Items.Clear();
            listBox12.Items.Clear();
            observaciones.Items.Clear();
            listBox20.Items.Clear();
            
            label3.Text = "Creando PDF";


        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)

        {
            
           
           
            if (progressBar1.Value==100)
            {
               
                label3.Text = "Archivos ya creados.";
            }
        }


    public void BuscarNombreMateria(string codmateria,string codespe, ref string nombremateria,ref string nombcarre)
        {
            
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                String strFila = Row.Index.ToString();
                string Valor = Convert.ToString(Row.Cells["Esp#"].Value);
                string Valor2 = Convert.ToString(Row.Cells["Materia"].Value);

                if (Valor == codespe)
                {
                    if (Valor2 == codmateria)
                    {
                        nombremateria=  Convert.ToString(Row.Cells[4].Value);
                        nombcarre = Convert.ToString(Row.Cells[1].Value);
                    }
                }
            }

            
            
        }


        private void LLenarGrid(string archivo, string hoja)
        {
            //declaramos las variables         
            OleDbConnection conexion = null;
            DataSet dataSet = null;
            OleDbDataAdapter dataAdapter = null;
            string consultaHojaExcel = "Select * from [" + hoja + "$]";

            //esta cadena es para archivos excel 2007 y 2010
            // string cadenaConexionArchivoExcel = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + archivo + "';Extended Properties=Excel 12.0;";
            //    string cadenaConexionArchivoExcel = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + archivo + "';Excel 12.0 xml;HDR=No;IMEX=1;Readonly=True;";
            //para archivos de 97-2003 usar la siguiente cadena



            string cadenaConexionArchivoExcel = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + archivo + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";


            //  conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";


            //Validamos que el usuario ingrese el nombre de la hoja del archivo de excel a leer
            if (string.IsNullOrEmpty(hoja))
            {
                MessageBox.Show("No hay una hoja para leer");
            }
            else
            {
                try
                {
                    //Si el usuario escribio el nombre de la hoja se procedera con la busqueda
                    conexion = new OleDbConnection(cadenaConexionArchivoExcel);//creamos la conexion con la hoja de excel
                    conexion.Open(); //abrimos la conexion
                    dataAdapter = new OleDbDataAdapter(consultaHojaExcel, conexion); //traemos los datos de la hoja y las guardamos en un dataSdapter
                    dataSet = new DataSet(); // creamos la instancia del objeto DataSet
                    dataAdapter.Fill(dataSet, hoja);//llenamos el dataset
                    dataGridView1.DataSource = dataSet.Tables[0]; //le asignamos al DataGridView el contenido del dataSet
                    conexion.Close();//cerramos la conexion
                    dataGridView1.AllowUserToAddRows = false;       //eliminamos la ultima fila del datagridview que se autoagrega
                }
                catch (Exception ex)
                {
                    //en caso de haber una excepcion que nos mande un mensaje de error
                    MessageBox.Show("Error, Verificar que el archivo sea el de los Datos de las Carreras.", ex.Message);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //creamos un objeto OpenDialog que es un cuadro de dialogo para buscar archivos
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Archivos de Excel (*.xls;*.xlsx)|*.xls;*.xlsx"; //le indicamos el tipo de filtro en este caso que busque
                                                                             //solo los archivos excel

            dialog.Title = "Seleccione el archivo de Excel";//le damos un titulo a la ventana

            dialog.FileName = string.Empty;//inicializamos con vacio el nombre del archivo

            //si al seleccionar el archivo damos Ok
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //el nombre del archivo sera asignado al textbox

                string hoja = "Todas"; //la variable hoja tendra el valor del textbox donde colocamos el nombre de la hoja
                LLenarGrid(dialog.FileName, hoja); //se manda a llamar al metodo
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; //se ajustan las
            }
        }

        private void chart1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog dlCarpeta = new FolderBrowserDialog();
            dlCarpeta.RootFolder = System.Environment.SpecialFolder.Desktop;
            dlCarpeta.ShowNewFolderButton = false;
            dlCarpeta.Description = "Selecciona la carpeta";
            if (dlCarpeta.ShowDialog() == DialogResult.OK)
            {
                txt_direccarpe.Text = dlCarpeta.SelectedPath+@"\";
            }
        }
    }
}
