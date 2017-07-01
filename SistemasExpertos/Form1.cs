using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SistemasExpertos
{
    public partial class Form1 : Form
    {
        List<Fiscal> Fiscalia1 = new List<Fiscal>();
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;//cambia el icono del puntero para indicar que se esta arrastrando un archivo
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);//se obtiene la ruta del archivo que se arrastro al datagridview
            txtXml1.Text = files[0];//colocamos la ruta del documento xml en el textbox 
            XDocument xmlDoc1 = XDocument.Load(@"" + files[0]);//a traves de la ruta obtenemos los datos del documento xmñ

            string nombreXml = Path.GetFileNameWithoutExtension(files[0]);//tambien necesitamos el nombre del documento sin su extension

            Fiscalia1 = xmlDoc1.Descendants( nombreXml+ "_row")//cada elemento grande del documento xml se llama igual que el nombre del documento mas la palabra _row
                                           .Select(c=>new Fiscal
                                           {
                                                Id_fiscal = (string)c.Element("id_fiscal"),
                                                Apell = (string)c.Element("apell"),
                                                Nomb = (string)c.Element("nomb"),
                                                Pend = (double)c.Element("pend"),
                                                Calif = (double)c.Element("calif"),
                                                Preli = (double)c.Element("preli"),
                                                Prepa = (double)c.Element("prepa"),
                                                Inter = (double)c.Element("inter"),
                                                Juzga = (double)c.Element("juzga"),
                                                Deriv = (double)c.Element("deriv"),
                                                Archi = (double)c.Element("archi"),
                                                Archicon = (double)c.Element("archicon"),
                                                Pnp = (double)c.Element("pnp"),
                                                Prov = (double)c.Element("prov"),
                                                Acum = (double)c.Element("acum"),
                                                Rpnp = (double)c.Element("rpnp"),
                                                Asig = (double)c.Element("asig"),
                                                Expe = (double)c.Element("expe"),
                                                Iexpe = (double)c.Element("iexpe"),
                                                Apel = (double)c.Element("apel"),
                                                Impug = (double)c.Element("impug"),
                                                Cons = (double)c.Element("cons"),
                                                Excl = (double)c.Element("excl"),
                                                Senten = (double)c.Element("senten"),
                                                Quejas = (double)c.Element("quejas"),
                                                Ipre = (double)c.Element("ipre"),
                                                Sobreseim = (double)c.Element("sobreseim"),
                                                Sobreseimcon = (double)c.Element("sobreseimcon"),
                                                Acuerep = (double)c.Element("acuerep"),
                                                Princoport = (double)c.Element("princoport"),
                                                Id_depen_mpub = (double)c.Element("id_depen_mpub"),
                                                Id_esp = (string)c.Element("id_esp"),
                                                Pdom = (double)c.Element("pdom"),
                                                Qpdom = (double)c.Element("qpdom")
                                            }).ToList();
            var source = new BindingSource();
            source.DataSource = Fiscalia1;
            dataGridView1.DataSource = source; 
        }

        private void dataGridView2_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dataGridView2_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtXml2.Text = files[0];
            
            DataSet dsXmlFile = new DataSet();
            dsXmlFile.ReadXml(@"" + files[0], XmlReadMode.Auto);
            
            dataGridView2.DataSource = dsXmlFile.Tables[0];
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            Reporte.Generar(Fiscalia1);
        }
    }
}
