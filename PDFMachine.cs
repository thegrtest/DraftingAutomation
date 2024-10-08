using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksToPdf
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // Use a FolderBrowserDialog to allow the user to select a directory
            string directoryPath = SelectDirectory();
            if (string.IsNullOrEmpty(directoryPath))
            {
                Console.WriteLine("No directory selected. Exiting the program.");
                return;
            }

            SldWorks swApp = new SldWorks();
            swApp.Visible = false;  // Make SolidWorks invisible
            swApp.FrameState = (int)swWindowState_e.swWindowMinimized;  // Minimize the SolidWorks window

            // Convert all .slddrw files in the directory to PDFs
            foreach (string filePath in Directory.GetFiles(directoryPath, "*.slddrw"))
            {
                ConvertSlddrwToPdf(swApp, filePath);
            }

            // Convert all .tif files in the directory to PDFs
            foreach (string filePath in Directory.GetFiles(directoryPath, "*.tif"))
            {
                ConvertTifToPdf(filePath);
            }

            swApp.ExitApp();
            swApp = null;
            GC.Collect();
        }

        static void ConvertSlddrwToPdf(SldWorks swApp, string filePath)
        {
            ModelDoc2 drawingDoc = (ModelDoc2)swApp.OpenDoc(filePath, (int)swDocumentTypes_e.swDocDRAWING);

            if (drawingDoc != null)
            {
                string pdfFilePath = Path.ChangeExtension(filePath, ".pdf");
                int errors = 0;
                int warnings = 0;

                bool saveResult = drawingDoc.Extension.SaveAs(pdfFilePath,
                                                             (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                                             (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                                             null,
                                                             ref errors,
                                                             ref warnings);

                // Properly close the document using CloseDoc
                swApp.CloseDoc(drawingDoc.GetTitle());

                if (saveResult && errors == 0)
                {
                    Console.WriteLine($"Successfully converted {filePath} to PDF.");
                }
                else
                {
                    Console.WriteLine($"Failed to convert {filePath} to PDF. Errors: {errors}, Warnings: {warnings}");
                }
            }
            else
            {
                Console.WriteLine($"Failed to open {filePath}. Ensure SolidWorks is installed and the file is accessible.");
            }
        }

        static void ConvertTifToPdf(string filePath)
        {
            string pdfFilePath = Path.ChangeExtension(filePath, ".pdf");

            using (PdfDocument document = new PdfDocument())
            {
                PdfPage page = document.AddPage();
                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    using (XImage image = XImage.FromFile(filePath))
                    {
                        page.Width = image.PixelWidth * 72 / image.HorizontalResolution;
                        page.Height = image.PixelHeight * 72 / image.VerticalResolution;
                        gfx.DrawImage(image, 0, 0, page.Width, page.Height);
                    }
                }
                document.Save(pdfFilePath);
            }

            Console.WriteLine($"Successfully converted {filePath} to PDF.");
        }

        // Method to open a FolderBrowserDialog and return the selected directory path
        static string SelectDirectory()
        {
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                folderBrowser.Description = "Select the directory containing .slddrw and .tif files";
                folderBrowser.ShowNewFolderButton = false;

                DialogResult result = folderBrowser.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                {
                    return folderBrowser.SelectedPath;
                }
                return null;
            }
        }
    }
}
