using iText.Kernel.Pdf.Filespec;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using iText.Kernel.Geom;
using System.Windows.Media.Animation;
using Microsoft.Win32;

namespace ETDA_Attach_PDF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()

        {
            InitializeComponent();

        }



        private void BtnSelectPDF_Click(object sender, MouseButtonEventArgs e)
        {
            
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".txt"; // Default file extension
            dialog.Filter = "PDF documents (.pdf)|*.pdf"; // Filter files by extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                string filename = dialog.FileName;
                Lb_PDFpath.Content = filename;
                Storyboard sb = FindResource("story_s1") as Storyboard;
                sb.Begin();

            }
        }

        private void BtnSavePath_Click(object sender, MouseButtonEventArgs e)
        {
            // Configure save file dialog box
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.FileName = "AttachPDF"; // Default file name
            dialog.DefaultExt = ".pdf"; // Default file extension
            dialog.Filter = "PDF documents (.pdf)|*.pdf"; // Filter files by extension

            // Show save file dialog box
            bool? result = dialog.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dialog.FileName;
                Lb_SavePath.Content = filename;
                Storyboard sb = FindResource("story_s3") as Storyboard;
                sb.Begin();
            }
        }

        private void BtnSelectAttach_Click(object sender, MouseButtonEventArgs e)
        {
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Multiselect = true;
                openFileDialog.Filter = "Attach documents (.pdf,jpg,tif,png)|*.pdf; *.tif; *.jpg; *.png";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (openFileDialog.ShowDialog() == true)
                {
                    foreach (string filename in openFileDialog.FileNames)
                        listb.Items.Add(filename);
                    Storyboard sb = FindResource("story_s2") as Storyboard;
                    sb.Begin();
                }
            }
        }

        private void BtnStart_Click(object sender, MouseButtonEventArgs e)
        {
            Storyboard sb = FindResource("reset") as Storyboard;
            try
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(Lb_PDFpath.Content.ToString()), new PdfWriter(Lb_SavePath.Content.ToString()));
                foreach (var lishBoxItem in listb.Items)
                {
                    var onlyFileName = System.IO.Path.GetFileName(lishBoxItem.ToString());
                    String embeddedFileName = onlyFileName;
                    String embeddedFileDescription = "Attachment";
                    byte[] embeddedFileContentBytes = File.ReadAllBytes(lishBoxItem.ToString()); ;


                    PdfFileSpec spec = PdfFileSpec.CreateEmbeddedFileSpec(pdfDoc, embeddedFileContentBytes,
                        embeddedFileDescription, embeddedFileName, null, null, null);

                    // This method adds file attachment at document level.
                    pdfDoc.AddFileAttachment(onlyFileName, spec);

                }
                pdfDoc.Close();
            }
            catch   (Exception ex)
            {
                MessageBox.Show("Error Occur",ex.ToString());
                sb.Begin();
            }
            
            MessageBox.Show("PDF Attached successful");
            
            sb.Begin();
        }
    }
}
