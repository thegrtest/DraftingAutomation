using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ImageMagick;

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

            // SolidWorks conversion for .slddrw files
            SldWorks swApp = new SldWorks();
            swApp.Visible = false;
            swApp.FrameState = (int)SolidWorks.Interop.swconst.swWindowState_e.swWindowMinimized;

            // Convert .slddrw files to PDFs
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.slddrw"))
            {
                ConvertSlddrwToPdf(swApp, filePath, outputDirectoryPath);
            }

            // Convert .tif files to PDFs
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.tif"))
            {
                ConvertTifToPdf(filePath, outputDirectoryPath);
            }

            // Convert .dwg files to PDFs using AutoCAD COM via late binding
            foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.dwg"))
            {
                ConvertDwgToPdfLateBinding(filePath, outputDirectoryPath);
            }

            swApp.ExitApp();
            swApp = null;

            GC.Collect();
        }

        static void ConvertSlddrwToPdf(SldWorks swApp, string filePath, string outputDirectory)
        {
            ModelDoc2 drawingDoc = (ModelDoc2)swApp.OpenDoc(filePath, (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocDRAWING);

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
            }

            Console.WriteLine($"Successfully converted {filePath} to {pdfFilePath}");
        }

        static void ConvertDwgToPdfLateBinding(string filePath, string outputDirectory)
        {
            object acadApp = null;
            try
            {
                // Create the AutoCAD application object using late binding
                Type acadType = Type.GetTypeFromProgID("AutoCAD.Application");
                if (acadType == null)
                {
                    Console.WriteLine("AutoCAD is not installed on this machine.");
                    return;
                }
                acadApp = Activator.CreateInstance(acadType);
                acadType.InvokeMember("Visible", BindingFlags.SetProperty, null, acadApp, new object[] { false });

                // Get the Documents collection
                object documents = acadType.InvokeMember("Documents", BindingFlags.GetProperty, null, acadApp, null);

                // Open the DWG file
                Type docType = documents.GetType();
                object acadDoc = docType.InvokeMember("Open", BindingFlags.InvokeMethod, null, documents, new object[] { filePath, false });

                // Get the Plot property from the document
                object plot = acadDoc.GetType().InvokeMember("Plot", BindingFlags.GetProperty, null, acadDoc, null);

                // Set up the PDF export options
                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                // Use late binding to call PlotToFile
                plot.GetType().InvokeMember("PlotToFile", BindingFlags.InvokeMethod, null, plot, new object[] { pdfFilePath, "DWG To PDF.pc3" });

                // Close the document without saving changes
                acadDoc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, acadDoc, new object[] { false });

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
                    acadApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, acadApp, null);
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
