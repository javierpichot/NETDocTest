using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace TestDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;

            // The Add method has four reference parameters, all of which are
            // optional. Visual C# allows you to omit arguments for them if
            // the default values are what you want.
            wordApp.Documents.Open("c:\\temp\\CVJavierPichot.docx");
            object missing = System.Reflection.Missing.Value;

            object findText = "Pichot";

            wordApp.Selection.Find.ClearFormatting();

            if (wordApp.Selection.Find.Execute(ref findText, missing, missing, missing, missing, missing, missing, missing,
                missing, "Reemplazo"))
            {
                MessageBox.Show("Text found.");
            }
            else
            {
                MessageBox.Show("The text could not be located.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;

            object template = "c:\\temp\\docs\\template.dotx";
            object toFile = "";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "doc files (*.doc)|*.doc|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //object toFile = "c:\\temp\\docs\\ReportFinish.doc";
                toFile =  saveFileDialog1.FileName ;
            }

            object missing = System.Reflection.Missing.Value;

            doc = app.Documents.Open(template, missing, missing);
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            app.Selection.Find.Execute("&EMPRESA&", missing, missing, missing, missing, missing, missing, missing,
                missing, txtRazonSocial.Text.ToString());
            app.Selection.Collapse();

            app.Selection.Find.Execute("&CUIT&", missing, missing, missing, missing, missing, missing, missing,
                missing, txtCUIT.Text.ToString());
            app.Selection.Collapse();

            app.Selection.Find.Execute("&MASASALARIAL&", missing, missing, missing, missing, missing, missing, missing,
                missing, txtMasaSalarial.Text.ToString());
            app.Selection.Collapse();

            object Save = (object)toFile;
            doc.SaveAs(Save, missing, missing, missing);

            doc.Close(false, missing, missing);
            app.Quit(false, false, false);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

            //Lo abro nuevamente!
            app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            app.Documents.Open(toFile);

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
