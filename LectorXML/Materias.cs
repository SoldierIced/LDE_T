using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
//using System.Xml.Linq;
using System.Collections;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;

namespace LectorXML
{
    public class PDFcreater


    {
        public int cont;
        public int x, y,pag;
        public string CadenaGuardar;
        public Document doc;
        public int conta=0;
        

        public iTextSharp.text.Font fontHeader_1 = FontFactory.GetFont("Calibri", 20, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC, new iTextSharp.text.BaseColor(36, 64, 97));
        public iTextSharp.text.Font fontHeader_11 = FontFactory.GetFont("Calibri", 15, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC, new iTextSharp.text.BaseColor(36, 64, 97));
        public iTextSharp.text.Font fontHeader_2 = FontFactory.GetFont("Bell MT", 15, iTextSharp.text.Font.BOLD, new iTextSharp.text.BaseColor(125, 125, 125));
        public iTextSharp.text.Font fontHeader_3 = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC, new iTextSharp.text.BaseColor(36, 64, 97));
        public iTextSharp.text.Font fontHeader_4 = FontFactory.GetFont("Arial", 12, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC, new iTextSharp.text.BaseColor(0,0, 0));
        public iTextSharp.text.Font fontHeader_44 = FontFactory.GetFont("Arial", 9, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC, new iTextSharp.text.BaseColor(0,0, 0));

        public static iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance("Image1.jpg");

        public string Titulogrande1;
        public string subtitulo;
        public string textoAagregar;
        public float marginLeft = 72;
        public float marginRight = 36;
        public float marginTop = 20;
        public float marginBottom = 50;
        PdfWriter writer;
        
        
        
        public void CargaP(string nombrearchivo,Materias m,string nombremateria,string nombrecarrera,int respuestas)
        {
           // fontHeader_3.SetStyle(iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC);
            


            doc = new Document(PageSize.LETTER);
            // el chart1 tiene que recibir

            // Creamos el documento con el tamaño de página tradicional
            y =545;
            x =10;
            pag = 0;

            cont = 1;
            string especialidad =m.esp ;
            string anio = m.anio ;
            string Curso = m.curso;
           
            string Titulogrande1 = "Encuesta de : " + nombremateria;
        
            string espe =especialidad;
            string aniot = anio;
            string curso = Curso;
        


            iTextSharp.text.Rectangle pageType = iTextSharp.text.PageSize.A4;
            float marginLeft = 10;
            float marginRight = 36;
            float marginTop = 20;
            float marginBottom = 50;

            try { CadenaGuardar += nombremateria + " " + m.curso + ".pdf"; }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
            }
            
            doc = new Document(pageType, marginLeft, marginRight, marginTop, marginBottom);
            

            // Indicamos donde vamos a guardar el documento
             writer = PdfWriter.GetInstance(doc,
                                        new FileStream(CadenaGuardar, FileMode.Create));
            doc.Open();

            //propiedades del pdf. ---------------------------------------
            Paragraph para1 = new Paragraph();

            img.ScaleToFit(40, 30);

            //Imagen - Esquina inferior izquierda
            img.SetAbsolutePosition(55,795);
            doc.Add(img);


                 
            Chunk cabezera = new Chunk("                              Ministerio de Educación, Universidad Tecnológica Nacional,  Facultad Regional General Pacheco.  ", fontHeader_44);
            para1.Add(cabezera);
            para1.Alignment = Element.ALIGN_MIDDLE;
            doc.Add(para1);
            // el chart1 tiene que recibir
            doc.AddTitle(Titulogrande1);//subtitulo
            doc.AddSubject(subtitulo);
            doc.AddKeywords("");
            doc.AddCreator("Graficador de Encuestas");
            doc.AddAuthor("Nehuen Fortes");
            doc.AddHeader("Owner", "Graficos de encuestas");


            Paragraph paraHeader_1 = new Paragraph(Titulogrande1, fontHeader_1);
            paraHeader_1.Alignment = Element.ALIGN_CENTER;
            paraHeader_1.SpacingAfter = 5f;
            doc.Add(paraHeader_1);

            //----------------------------------------------- subtitulo-----------
            // Nombre de la carrera: Tecnicatura Superior en Sistemas Informáticos
          //  Especialidad: 120       Año: 2017       Curso: 120 - 1C

                          Paragraph titulo = new Paragraph();
            titulo.Alignment = Element.ALIGN_CENTER;
            Chunk nombre = new Chunk("Nombre de la carrera: ", fontHeader_4);
            Chunk texto2 = new Chunk(nombrecarrera + '\n', fontHeader_3);
            titulo.Add(nombre);
            titulo.Add(texto2);
            Chunk esp_c = new Chunk("Especialidad: ", fontHeader_4);
            Chunk texto3 = new Chunk(especialidad, fontHeader_3);
            titulo.Add(esp_c);
            titulo.Add(texto3);
            esp_c = new Chunk("        Año: ", fontHeader_4);
             texto3 = new Chunk(anio , fontHeader_3);
            titulo.Add(esp_c);
            titulo.Add(texto3);
            esp_c = new Chunk("        Curso: ", fontHeader_4);
            texto3 = new Chunk(curso + '\n', fontHeader_3);
            titulo.Add(esp_c);
            titulo.Add(texto3);
            esp_c = new Chunk("Total de Alumnos: ", fontHeader_4);
            texto3 = new Chunk(respuestas.ToString(), fontHeader_3);
            titulo.Add(esp_c);
            titulo.Add(texto3);
            doc.Add(titulo);
            //----------------------------------------------------------------------------
        }

        public void Guardar(Chart chart1 )
        {
            pag++;
        var chartimage = new MemoryStream();
         
        chart1.SaveImage(chartimage, ChartImageFormat.Png);

            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());


            //  Chart_image.ScalePercent(98.6f);
            // iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage);// 'Dirreccion a la imagen que se hace referencia
            

            Chart_image.SetAbsolutePosition(x,y);// 'Posicion en el eje cartesiano
           
           Chart_image.ScaleAbsoluteWidth(270);// 'Ancho de la imagen
            Chart_image.ScaleAbsoluteHeight(160);// 'Altura de la imagen
            doc.Add(Chart_image);// ' Agrega la imagen al documento

            if (pag ==8) { doc.NewPage();
                iTextSharp.text.Rectangle pageType = iTextSharp.text.PageSize.A4;
                PdfContentByte cb = writer.DirectContent;
                cb.MoveTo(marginLeft, marginTop);
                cb.LineTo(1, marginTop);
                cb.Stroke();
           /*
               
            */
                x = 10; y = 600; pag = 0;cont = 0; }

                    }

        public void agregarObser (string obse,int preg)
        {

            Paragraph a = new Paragraph("\n");
            a.Alignment = Element.ALIGN_RIGHT;
            a.SpacingAfter = 1f;


            Paragraph observacion= new Paragraph("", fontHeader_4); 
            observacion.Alignment = Element.ALIGN_LEFT;
            observacion.SpacingAfter = 1f;

            if (conta==0) {
                for (int h = 0; h <45; h++) {                                     
            doc.Add(a);
                }
                conta++;
            }

            if (preg == 1)
            {
                 observacion = new Paragraph(obse, fontHeader_11);
                
            }
            else {
                
                observacion = new Paragraph(obse);

            }
            if (obse != string.Empty) { 
                doc.Add(observacion);
            }

        }

        public void CerrarPDF()
        {

            doc.Close();

        }


}


  


     public class Materias
    {
       public string anio, esp, COD;
       public  string curso;
        public int ejey,ejex;
        public int [,] m;
       public  int cantpreg;
        public string[] obs;
        public string[] preg;
        public int respuestas;


        public  Materias (int cantidadPreguntas ,int cantidadrespuestasP)
        {
            m = new int[cantidadPreguntas, cantidadrespuestasP];

            anio = esp =curso= COD = "";

            ejex = cantidadPreguntas;
            ejey = cantidadrespuestasP;

            obs = new string[cantidadPreguntas];
            preg = new string[cantidadPreguntas];
            cantpreg = cantidadPreguntas;
        }
       public void CargarObser  (string observacion, int pregunta,string respuesta,string preguntaS)
        {

            obs[pregunta] +=  " Respuesta: " + respuesta + "\n"+
               
                 " Observacion: " + observacion+ "\n" +"------------------------------------------" + "\n" ;
            preg[pregunta] = preguntaS;
        }
       
    }
}
