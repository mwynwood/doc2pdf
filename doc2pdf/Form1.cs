using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace doc2pdf
{

    public partial class Form1 : Form
    {
        public void generateCoverPage(string outputFileName)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created with PDFsharp";

            // Create an empty page
            PdfPage page = document.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Verdana", 20, XFontStyle.BoldItalic);

            // Draw the text
            gfx.DrawString("Hello, World!", font, XBrushes.Black,
            new XRect(0, 0, page.Width, page.Height),
            XStringFormats.Center);

            // Save the document...
            document.Save(outputFileName);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ButtonDocs_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowNewFolderButton = false;

            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxDocs.Text = folderBrowserDialog1.SelectedPath;
            }            
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
            saveFileDialog1.ShowDialog();
            generateCoverPage(saveFileDialog1.FileName);
        }
    }
}
