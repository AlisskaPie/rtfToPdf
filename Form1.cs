using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using Xceed.Words.NET;


namespace rtfToDocx
{
    public partial class Form1 : Form
    {
        Word._Application word_app = new Word.ApplicationClass();

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                checkBox1.Checked = false;
                textBox2.Text = folderDlg.SelectedPath;
                Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox2.Text = textBox1.Text;
            }
        }

        /* public static string[] GetFiles(string path, string searchPattern, SearchOption searchOption)
         {
             string[] searchPatterns = searchPattern.Split('|');
             List<string> files = new List<string>();
             foreach (string sp in searchPatterns)
                 files.AddRange(System.IO.Directory.GetFiles(path, sp, searchOption));
             files.Sort();
             return files.ToArray();
         }*/




        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            // Make Word visible (optional).
            // word_app.Visible = true;
            object format_doc = (int)16;
            // Open the file.
            var files = System.IO.Directory.GetFiles(textBox1.Text, "*.rtf");
            progressBar1.Maximum = files.Length;

            string workingDir = Path.GetFullPath(textBox2.Text);
            string docxFilePath = Path.Combine(workingDir, @"SAR.docx");

            //DocumentCore doc = new DocumentCore();
            using (DocX dc = DocX.Create(docxFilePath))
            {
                foreach (var f in files)
                {
                    object input_file = f;
                    // Save the output file.
                    object output_file = f + ".docx";
                    var d = word_app.Documents.Open(ref input_file);
                        d.SaveAs(ref output_file, ref format_doc);
                    d.Close();

                    // Exit the server without prompting.
                    object false_obj = false;
                    progressBar1.PerformStep();
                    //doc = DocumentCore.Load(output_file.ToString());
                    //dc.InsertDocument(docx, true);

                    using (DocX oldDocument = DocX.Load(output_file.ToString()))
                    {
                        //oldDocument.InsertSection();
                        //var p = oldDocument.InsertParagraph("My Heading");
                        //p.StyleName = "Heading3";

                        //Paragraph p1 = oldDocument.InsertParagraph("", false);
                        // p1.InsertPageBreakAfterSelf();

                        //string pageBreak = "\f";
                        Paragraph p3 = oldDocument.Paragraphs.ElementAt(2);
                        p3.StyleName = "Heading3";
                        oldDocument.Paragraphs.Last().InsertPageBreakAfterSelf();

                        //oldDocument.GetSections().Last().SectionBreakType = SectionBreakType.defaultNextPage;

                        dc.InsertDocument(oldDocument);
                    }
                }
                dc.Paragraphs.Last().Remove(false);
                dc.Save();
            }
            
            MessageBox.Show("Done");
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

