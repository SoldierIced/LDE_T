using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LectorXML
{
    public partial class DatosMateria : Form
    {
        public DatosMateria( Materias m,ListBox preguntas, ListBox respuestas)
        {
            textBox1.Text = "Materia : " + m.COD;

            int i = 0;
            textBox2.Text = "Numero pregunta: " + i + "texto de la pregunta: " + preguntas.Items[i].ToString();
            // currencyManager = (CurrencyManager)dataGrid1.BindingContext[m];
            InitializeComponent();
    }

        private void label2_Click(object sender, EventArgs e)
        {
          
        }

        private void DatosMateria_Load(object sender, EventArgs e)
        {

        }
    }
}
