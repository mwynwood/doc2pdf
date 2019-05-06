using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Collections;

using Microsoft.Office.Interop;

using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Drawing;

namespace doc2pdf
{

    public partial class Form1 : Form 
    {
        ArrayList allTheFiles = new ArrayList(); // Used to keep track of the DOCs
        ArrayList allThePDFs = new ArrayList(); // Used to keep a track of the PDFs
        string coverPageFileName = "CoverPage.pdf";

        public void generateCoverPage(string outputFileName, string theImage)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "doc2pdf Cover Page";

            // Create an empty page
            PdfPage page = document.AddPage();
            //page.Width = 595
            //page.Height = 842

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // Create a font
            XFont font = new XFont("Calibri", 20, XFontStyle.Bold);

            // Draw the image
            if (theImage.Length > 0)
            {
                XImage image = XImage.FromFile(theImage);
                gfx.DrawImage(image, 200, 200, 200, 200); // this makes it square. scale it!
            }
            
            // Draw the text
            gfx.DrawString(textBoxLine1.Text, font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine2.Text, font, XBrushes.Black, new XRect(0, 40, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine3.Text, font, XBrushes.Black, new XRect(0, 80, page.Width, page.Height), XStringFormats.Center);
            gfx.DrawString(textBoxLine4.Text, font, XBrushes.Black, new XRect(0, 120, page.Width, page.Height), XStringFormats.Center);

            // Save the document...
            document.Save(outputFileName);
            document.Close();
        }

        public FileInfo convertDoc2Pdf(FileInfo documentToConvert)
        {
            var appWord = new Microsoft.Office.Interop.Word.Application();

            if (documentToConvert.Extension == ".pdf" && documentToConvert.Name != coverPageFileName)
            {
                // If it's already a PDF, just copy it to the working dir
                documentToConvert.CopyTo(Application.StartupPath + "\\" + documentToConvert.Name, true);
                return new FileInfo(Application.StartupPath + "\\" + documentToConvert.Name);
            }else if (appWord.Documents != null)
            {
                try
                {
                    var wordDocument = appWord.Documents.Open(documentToConvert.FullName);
                    FileInfo pdfDocName = new FileInfo(Application.StartupPath + "\\" + documentToConvert.Name + ".pdf");

                    if (wordDocument != null)
                    {
                        wordDocument.SaveAs2(pdfDocName.FullName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                        wordDocument.Close();
                        appWord.Quit();
                        return pdfDocName;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            appWord.Quit();
            return null;
        }

        public Form1()
        {
            InitializeComponent();
            this.Text = Application.ProductName;
            this.Icon = doc2pdf.Properties.Resources.merge_ico;

            FileInfo fiCoverPage = new FileInfo(coverPageFileName);
            allTheFiles.Insert(0, fiCoverPage);

            updateInterface();
        }

        public void updateInterface()
        {
            listBoxDocs.DataSource = null;
            listBoxDocs.DataSource = allTheFiles;
            listBoxDocs.DisplayMember = "Name";

            if(allTheFiles.Count == 0)
            {
                buttonDocRemove.Enabled = false;
                buttonDocMoveUp.Enabled = false;
                buttonDocMoveDown.Enabled = false;
                buttonGo.Enabled = false;
            } else
            {
                buttonDocRemove.Enabled = true;
                buttonDocMoveUp.Enabled = true;
                buttonDocMoveDown.Enabled = true;
                buttonGo.Enabled = true;
            }

            if (allTheFiles.Count <= 2 && checkBoxCoverPage.Checked == true)
            {
                buttonDocMoveUp.Enabled = false;
                buttonDocMoveDown.Enabled = false;
            }
            else if (allTheFiles.Count <= 1 && checkBoxCoverPage.Checked == false)
            {
                buttonDocMoveUp.Enabled = false;
                buttonDocMoveDown.Enabled = false;
            }
            else { 
                buttonDocMoveUp.Enabled = true;
                buttonDocMoveDown.Enabled = true;
            }

            if (textBoxLogo.Text != "")
            {
                pictureBoxLogo.ImageLocation = textBoxLogo.Text;
            } else
            {
                pictureBoxLogo.Image = doc2pdf.Properties.Resources.merge_png;
            }
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
            saveFileDialog1.FileName = "merged";
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString();

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor.Current = Cursors.WaitCursor;

                // Generate the Cover Page
                if (checkBoxCoverPage.Checked == true)
                {
                    generateCoverPage(coverPageFileName, textBoxLogo.Text);
                }

                // Convert all the DOC and DOCX files to PDFs
                // And put the PDFs in a new ArrayList
                allThePDFs.Clear();
                foreach (FileInfo doc in allTheFiles)
                {
                    allThePDFs.Add(convertDoc2Pdf(doc));
                }

                // Merge the PDF files into one
                try
                {
                    PdfDocument outputDocument = new PdfDocument();
                    foreach (FileInfo file in allThePDFs)
                    {
                        // Attention: must be in Import mode
                        var mode = PdfDocumentOpenMode.Import;
                        var inputDocument = PdfReader.Open(file.FullName, mode);
                        int totalPages = inputDocument.PageCount;
                        for (int pageNo = 0; pageNo < totalPages; pageNo++)
                        {
                            // Get the page from the input document...
                            PdfPage page = inputDocument.Pages[pageNo];
                            // ...and copy it to the output document.
                            outputDocument.AddPage(page);
                        }
                    }
                    // Save the document
                    outputDocument.Save(saveFileDialog1.FileName);
                    outputDocument.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                // Delete the PDF files
                if (checkBoxDeletePdfsAfterMerge.Checked == true)
                {
                    try
                    {
                        foreach (FileInfo file in allThePDFs)
                        {
                            file.Delete();
                        }
                        FileInfo cp = new FileInfo(coverPageFileName);
                        cp.Delete();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

                Cursor.Current = Cursors.Default;
                if (MessageBox.Show("Merge Complete!" + Environment.NewLine + "View the merged file?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                }
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
                FileInfo fiCoverPage = new FileInfo(coverPageFileName);
                allTheFiles.Insert(0, fiCoverPage);
            }
            updateInterface();
            listBoxDocs.ClearSelected();
        }

        private void ButtonDocAdd_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Documents|*.DOC;*.DOCX;*.PDF";
            openFileDialog1.Title = "Select a document";
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileInfo fi = new FileInfo(@openFileDialog1.FileName);
                allTheFiles.Insert(allTheFiles.Count, fi);
                updateInterface();
                listBoxDocs.SelectedIndex = allTheFiles.Count - 1;
            }
            
        }

        private void ButtonDocRemove_Click(object sender, EventArgs e)
        {
            if (listBoxDocs.SelectedIndex > -1)
            {
                if (allTheFiles[listBoxDocs.SelectedIndex].ToString() != coverPageFileName)
                {
                    allTheFiles.RemoveAt(listBoxDocs.SelectedIndex);
                    updateInterface();
                }
                else
                {
                    checkBoxCoverPage.Checked = false;
                }
            }
        }

        private void ButtonDocMoveUp_Click(object sender, EventArgs e)
        {
            int selection = listBoxDocs.SelectedIndex;
            int toppest = 0;

            if(checkBoxCoverPage.Checked == true)
            {
                toppest = 1;
            } else
            {
                toppest = 0;
            }

            if (selection > toppest)
            {
                Object selected = allTheFiles[selection];
                Object above = allTheFiles[selection - 1];

                allTheFiles.RemoveAt(selection);
                allTheFiles.RemoveAt(selection - 1);

                allTheFiles.Insert(selection - 1, selected);
                allTheFiles.Insert(selection, above);

                listBoxDocs.SelectedIndex = selection - 1;

                updateInterface();
            }            
        }

        private void ButtonDocMoveDown_Click(object sender, EventArgs e)
        {
            int selection = listBoxDocs.SelectedIndex;

            if(selection == allTheFiles.Count - 1)
            {
                //MessageBox.Show("already at the bottom");
            } else if(selection == 0 && checkBoxCoverPage.Checked == true)
            {
                //MessageBox.Show("you can't move the Cover Page down");
            } else
            {
                Object selected = allTheFiles[selection];
                Object below = allTheFiles[selection + 1];

                allTheFiles.RemoveAt(selection + 1);
                allTheFiles.RemoveAt(selection);

                allTheFiles.Insert(selection, below);
                allTheFiles.Insert(selection + 1, selected);

                listBoxDocs.SelectedIndex = selection + 1;

                updateInterface();
            }
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.ShowDialog();
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "doc2pdf Settings|*.doc2pdf";
            saveFileDialog1.Title = "Save Your Settings";
            saveFileDialog1.DefaultExt = ".doc2pdf";
            saveFileDialog1.FileName = "settings";
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString();

            if(saveFileDialog1.ShowDialog() == DialogResult.OK) { 
                try { 
                    using (StreamWriter writer = new StreamWriter(saveFileDialog1.FileName, false))
                    {
                        writer.WriteLine(checkBoxCoverPage.Checked.ToString());
                        writer.WriteLine(checkBoxDeletePdfsAfterMerge.Checked.ToString());
                        writer.WriteLine(textBoxLine1.Text);
                        writer.WriteLine(textBoxLine2.Text);
                        writer.WriteLine(textBoxLine3.Text);
                        writer.WriteLine(textBoxLine4.Text);
                        writer.WriteLine(textBoxLogo.Text);
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        writer.WriteLine("-"); // Spare spot for a future setting
                        foreach (Object obj in allTheFiles)
                        {
                            writer.WriteLine(obj.ToString());
                        }
                        writer.Close();
                        //MessageBox.Show("Settings Saved", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                } catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void OpenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "doc2pdf Settings|*.doc2pdf";
            openFileDialog1.Title = "Open Settings";
            openFileDialog1.DefaultExt = ".doc2pdf";
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string spareSetting;
                    TextReader tr;
                    tr = File.OpenText(openFileDialog1.FileName);
                    checkBoxCoverPage.Checked = bool.Parse(tr.ReadLine());
                    checkBoxDeletePdfsAfterMerge.Checked = bool.Parse(tr.ReadLine());
                    textBoxLine1.Text = tr.ReadLine();
                    textBoxLine2.Text = tr.ReadLine();
                    textBoxLine3.Text = tr.ReadLine();
                    textBoxLine4.Text = tr.ReadLine();
                    textBoxLogo.Text = tr.ReadLine();
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    spareSetting = tr.ReadLine();// Spare spot for a future setting
                    allTheFiles.Clear();
                    string line;
                    while ((line = tr.ReadLine()) != null)
                    {
                        allTheFiles.Add(new FileInfo(line));
                    }
                    tr.Close();
                    updateInterface();
                } catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void ButtonRemoveLogo_Click(object sender, EventArgs e)
        {
            textBoxLogo.Text = "";
            updateInterface();
        }
    }
}