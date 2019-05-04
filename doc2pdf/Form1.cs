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
using System.Collections;

using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace doc2pdf
{

    public partial class Form1 : Form 
    {

        ArrayList allTheFiles = new ArrayList();

        public void generateCoverPage(string outputFileName, string theImage)
        {
            PdfDocument document = new PdfDocument();
            //document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Calibri", 20, XFontStyle.Bold);

            // Draw the text
            gfx.DrawString(textBoxLine1.Text, font, XBrushes.Black, new XRect(0,   0, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine2.Text, font, XBrushes.Black, new XRect(0,  40, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine3.Text, font, XBrushes.Black, new XRect(0,  80, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine4.Text, font, XBrushes.Black, new XRect(0, 120, page.Width, page.Height), XStringFormats.Center);

            // Draw the image
            if (theImage.Length > 0)
            {
                XImage image = XImage.FromFile(theImage);
                gfx.DrawImage(image, 200, 200, 200, 200); // this makes it square. scale it!
            }

            // Save the document...
            document.Save(outputFileName);
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = Application.ProductName;

            //allTheFiles.Insert(0, "Cover Page");
            FileInfo fiCoverPage = new FileInfo("CoverPage.pdf");
            allTheFiles.Insert(0, fiCoverPage);

            updateListBox();
        }

        public void updateListBox()
        {
            listBoxDocs.DataSource = null;
            listBoxDocs.DataSource = allTheFiles;
            listBoxDocs.DisplayMember = "Name";
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }    

        private void ButtonLogo_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Image Files|*.BMP;*.JPG;*.GIF,*.PNG";
            openFileDialog1.Title = "Select a Logo for the Cover Page";
            openFileDialog1.FileName = "";
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxLogo.Text = openFileDialog1.FileName;
                pictureBoxLogo.ImageLocation = openFileDialog1.FileName;
            }
        }

        private void ButtonGo_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "PDF File|*.PDF";
            saveFileDialog1.Title = "Save the merged document";
            saveFileDialog1.DefaultExt = ".pdf";
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString();

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                generateCoverPage(saveFileDialog1.FileName, textBoxLogo.Text);
                MessageBox.Show("Done!", Application.ProductName);
            }
        }

        private void CheckBoxCoverPage_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBoxCoverPage.Checked == false)
            {
                groupBoxCoverPage.Enabled = false;
                allTheFiles.RemoveAt(0);
            } else
            {
                groupBoxCoverPage.Enabled = true;                
                //allTheFiles.Insert(0, "Cover Page");
                FileInfo fiCoverPage = new FileInfo("CoverPage.pdf");
                allTheFiles.Insert(0, fiCoverPage);
            }
            updateListBox();
        }

        private void ButtonDocAdd_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Documents|*.DOC;*.DOCX;*.PDF";
            openFileDialog1.Title = "Select a document";
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileInfo fi = new FileInfo(@openFileDialog1.FileName);
                //allTheFiles.Add(fi);
                allTheFiles.Insert(allTheFiles.Count, fi);
                updateListBox();
            }
            
        }

        private void ButtonDocRemove_Click(object sender, EventArgs e)
        {
            if (listBoxDocs.SelectedIndex > -1)
            {
                if (allTheFiles[listBoxDocs.SelectedIndex].ToString() != "CoverPage.pdf")
                {
                    allTheFiles.RemoveAt(listBoxDocs.SelectedIndex);
                    updateListBox();
                }
                else
                {
                    checkBoxCoverPage.Checked = false;
                }
            }
        }
    }
}
