using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ImageMagick;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Geometry;

namespace CombinedProcessor
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

            // Convert .slddrw files to PDFs
            SldWorks swApp = new SldWorks();
            swApp.Visible = false;  // Run SolidWorks in the background

            foreach (string filePath in Directory.GetFiles(directoryPath, "*.slddrw"))
            {
                ConvertSlddrwToPdf(swApp, filePath);
            }

            // Convert .tif files to PDFs
            foreach (string filePath in Directory.GetFiles(directoryPath, "*.tif"))
            {
                ConvertTifToPdf(filePath);
            }

            // Close SolidWorks
            swApp.ExitApp();
            swApp = null;

            // Extract text from bottom-right corner of PDFs
            foreach (string pdfFilePath in Directory.GetFiles(directoryPath, "*.pdf"))
            {
                ExtractTextFromPdfBottomRight(pdfFilePath);
            }
        }

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

        static void ConvertSlddrwToPdf(SldWorks swApp, string filePath)
        {
            ModelDoc2 drawingDoc = (ModelDoc2)swApp.OpenDoc(filePath, (int)swDocumentTypes_e.swDocDRAWING);

            if (drawingDoc != null)
            {
                string pdfFilePath = Path.ChangeExtension(filePath, ".pdf");
                drawingDoc.Extension.SaveAs(pdfFilePath,
                                            (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                            null, out int errors, out int warnings);
                swApp.CloseDoc(drawingDoc.GetTitle());
                Console.WriteLine($"Converted {filePath} to {pdfFilePath}");
            }
        }

        static void ConvertTifToPdf(string filePath)
        {
            string pdfFilePath = Path.ChangeExtension(filePath, ".pdf");

            using (PdfDocument document = new PdfDocument())
            {
                using (MagickImageCollection images = new MagickImageCollection(filePath))
                {
                    foreach (MagickImage image in images)
                    {
                        PdfPage page = document.AddPage();
                        page.Width = image.Width * 72 / image.Density.X;
                        page.Height = image.Height * 72 / image.Density.Y;

                        using (XGraphics gfx = XGraphics.FromPdfPage(page))
                        {
                            using (MemoryStream memoryStream = new MemoryStream())
                            {
                                image.Format = MagickFormat.Png;
                                image.Write(memoryStream);
                                memoryStream.Position = 0;

                                using (XImage xImage = XImage.FromStream(memoryStream))
                                {
                                    gfx.DrawImage(xImage, 0, 0, page.Width, page.Height);
                                }
                            }
                        }
                    }
                }
                document.Save(pdfFilePath);
                Console.WriteLine($"Converted {filePath} to {pdfFilePath}");
            }
        }

        static void ExtractTextFromPdfBottomRight(string pdfFilePath)
        {
            using (PdfDocument pdfDocument = PdfDocument.Open(pdfFilePath))
            {
                foreach (Page page in pdfDocument.GetPages())
                {
                    var width = page.Width;
                    var height = page.Height;

                    // Define the region to extract text from (bottom-right corner)
                    var bottomRightRegion = new PdfRectangle(
                        width - 150,
                        0,
                        width,
                        150
                    );

                    // Extract text from the defined region
                    var words = page.GetWords().Where(w =>
                        w.BoundingBox.BottomLeft.X >= bottomRightRegion.Left &&
                        w.BoundingBox.BottomLeft.Y >= bottomRightRegion.Bottom &&
                        w.BoundingBox.TopRight.X <= bottomRightRegion.Right &&
                        w.BoundingBox.TopRight.Y <= bottomRightRegion.Top
                    );

                    Console.WriteLine($"Text extracted from {pdfFilePath}:");
                    foreach (var word in words)
                    {
                        Console.WriteLine(word.Text);
                    }
                }
            }
        }
    }
}
