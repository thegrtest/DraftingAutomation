using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ImageMagick;
using Autodesk.AutoCAD.Interop;   // Add AutoCAD Interop for DWG handling
using Autodesk.AutoCAD.Interop.Common;
using SwConst;
using swDocumentTypes_e = SolidWorks.Interop.swconst.swDocumentTypes_e;

namespace SolidWorksToPdf
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // Use a FolderBrowserDialog to allow the user to select a directory for input files
            string inputDirectoryPath = SelectDirectory("Select the directory containing .slddrw, .tif, and .dwg files");
            if (string.IsNullOrEmpty(inputDirectoryPath))
            {
                Console.WriteLine("No input directory selected. Exiting the program.");
                return;
            }

            // Use a FolderBrowserDialog to allow the user to select a directory for saving PDFs
            string outputDirectoryPath = SelectDirectory("Select the directory where PDFs will be saved");
            if (string.IsNullOrEmpty(outputDirectoryPath))
            {
                Console.WriteLine("No output directory selected. Exiting the program.");
                return;
            }

            SldWorks swApp = new SldWorks();
            swApp.Visible = false;  // Make SolidWorks invisible
            swApp.FrameState = (int)SolidWorks.Interop.swconst.swWindowState_e.swWindowMinimized;  // Minimize the SolidWorks window

            // Convert all .slddrw files in the directory to PDFs
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.slddrw"))
            {
                ConvertSlddrwToPdf(swApp, filePath, outputDirectoryPath);
            }

            // Convert all .tif files in the directory to PDFs
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.tif"))
            {
                ConvertTifToPdf(filePath, outputDirectoryPath);
            }

            // Convert all .dwg files in the directory to PDFs using AutoCAD COM
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.dwg"))
            {
                ConvertDwgToPdf(filePath, outputDirectoryPath);
            }

            swApp.ExitApp();
            swApp = null;

            GC.Collect();
        }

        static void ConvertSlddrwToPdf(SldWorks swApp, string filePath, string outputDirectory)
        {
            ModelDoc2 drawingDoc = (ModelDoc2)swApp.OpenDoc(filePath, (int)swDocumentTypes_e.swDocDRAWING);

            if (drawingDoc != null)
            {
                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
                int errors = 0;
                int warnings = 0;

                bool saveResult = drawingDoc.Extension.SaveAs(pdfFilePath,
                                                             (int)SolidWorks.Interop.swconst.swSaveAsVersion_e.swSaveAsCurrentVersion,
                                                             (int)SolidWorks.Interop.swconst.swSaveAsOptions_e.swSaveAsOptions_Silent,
                                                             null,
                                                             ref errors,
                                                             ref warnings);

                // Properly close the document using CloseDoc
                swApp.CloseDoc(drawingDoc.GetTitle());

                if (saveResult && errors == 0)
                {
                    Console.WriteLine($"Successfully converted {filePath} to {pdfFilePath}");
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

        static void ConvertTifToPdf(string filePath, string outputDirectory)
        {
            string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

            using (PdfDocument document = new PdfDocument())  // Direct reference to PdfSharp's PdfDocument
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
            }

            Console.WriteLine($"Successfully converted {filePath} to {pdfFilePath}");
        }

        static void ConvertDwgToPdf(string filePath, string outputDirectory)
        {
            // Initialize AutoCAD
            AcadApplication acadApp = null;
            try
            {
                acadApp = (AcadApplication)Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application"), true);
                acadApp.Visible = false;

                // Open the DWG file
                AcadDocument acadDoc = acadApp.Documents.Open(filePath, false);

                // Set up the PDF export options
                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                acadDoc.Plot.PlotToFile(pdfFilePath, "DWG To PDF.pc3");

                acadDoc.Close(false);  // Close the document without saving changes
                Console.WriteLine($"Successfully converted {filePath} to {pdfFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting {filePath} to PDF: {ex.Message}");
            }
            finally
            {
                // Quit AutoCAD
                if (acadApp != null)
                {
                    acadApp.Quit();
                }
            }
        }

        static string SelectDirectory(string description)
        {
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                folderBrowser.Description = description;
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
