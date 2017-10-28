using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ReadPSMRequestForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = @"C:\SunXin\Competence Team Table\PSM request\TestForm.pdf";
         
        }
        void convertPDF(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {

                var dic = reader.AcroFields;
                StreamWriter sw = new StreamWriter("text.txt");
                foreach (var f in dic.Fields)
                {
                    sw.WriteLine("{0} : {1}", f.Key, dic.GetField(f.Key));
                }
                sw.Close();

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.RestoreDirectory = true;
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    foreach(var f in ofd.FileNames)
                    {
                        convertPDF(f);
                    }
                }
            }
        }
    }
}
