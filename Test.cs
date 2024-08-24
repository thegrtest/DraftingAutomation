using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ImageMagick;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;  // Alias PdfPig's PdfDocument
using PdfPigPage = UglyToad.PdfPig.Content.Page;    // Alias PdfPig's Page
using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;  // Alias PdfSharp's PdfDocument
using UglyToad.PdfPig.Geometry;  // Correct namespace for PdfRectangle

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

            // After generating all PDFs, extract text from the bottom right corner
            foreach (string pdfFilePath in Directory.GetFiles(directoryPath, "*.pdf"))
            {
                ExtractTextFromPdfBottomRight(pdfFilePath);
            }

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

            using (PdfSharpDocument document = new PdfSharpDocument())
            {
                using (MagickImageCollection images = new MagickImageCollection(filePath))
                {
                    foreach (MagickImage image in images)
                    {
                        PdfPage page = document.AddPage();
                        page.Width = image.Width * 72 / image.Density.X;
                        page.Height = image.Height * 72 / image.Density.Y;

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            // Convert MagickImage to a bitmap-compatible format (e.g., PNG)
                            image.Format = MagickFormat.Png;
                            image.Write(memoryStream);
                            memoryStream.Position = 0;

                            using (XImage xImage = XImage.FromStream(memoryStream))
                            {
                                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                                {
                                    gfx.DrawImage(xImage, 0, 0, page.Width, page.Height);
                                }
                            }
                        }
                    }
                }

                document.Save(pdfFilePath);
            }

            Console.WriteLine($"Successfully converted {filePath} to PDF.");
        }

        static void ExtractTextFromPdfBottomRight(string pdfFilePath)
        {
            using (PdfPigDocument pdfDocument = PdfPigDocument.Open(pdfFilePath))
            {
                foreach (PdfPigPage page in pdfDocument.GetPages())
                {
                    var width = page.Width;
                    var height = page.Height;

                    // Define the region to extract text from (bottom-right corner)
                    var bottomLeft = new PdfPoint(width - 150, 0);
                    var topRight = new PdfPoint(width, 150);
                    var bottomRightRegion = new PdfRectangle(bottomLeft, topRight);

                    // Extract text from the defined region
                    var words = page.GetWords().Where(w => bottomRightRegion.Contains(w.BoundingBox.BottomLeft)).ToList();

                    Console.WriteLine($"Extracted text from {pdfFilePath}:");
                    foreach (var word in words)
                    {
                        Console.WriteLine(word.Text);
                    }
                }
            }
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
