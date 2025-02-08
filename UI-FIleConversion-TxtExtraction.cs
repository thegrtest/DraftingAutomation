using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using ImageMagick;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

// Define Namespace Aliases to avoid conflicts
using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;
using PdfSharpPage = PdfSharp.Pdf.PdfPage;
using iTextPdfReader = iText.Kernel.Pdf.PdfReader;
using WinFormsApp1;

namespace SolidWorksToPdf
{
    class Program
    {
        static string logFile = "debug_log.txt";

        [STAThread]
        static void Main(string[] args)
        {
            AppContext.SetSwitch("System.Drawing.EnableUnixSupport", false);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1()); // Ensure the UI form starts
            Log("Program started.");

            string inputDirectoryPath = SelectDirectory("Select the directory containing .slddrw, .tif, and .dwg files");
            if (string.IsNullOrEmpty(inputDirectoryPath))
            {
                Log("ERROR: No input directory selected. Exiting program.");
                return;
            }
            Log($"✅ Input directory selected: {inputDirectoryPath}");

            string outputDirectoryPath = SelectDirectory("Select the directory where PDFs will be saved");
            if (string.IsNullOrEmpty(outputDirectoryPath))
            {
                Log("ERROR: No output directory selected. Exiting program.");
                return;
            }
            Log($"✅ Output directory selected: {outputDirectoryPath}");

            try
            {
                Log("Initializing SolidWorks...");
                Type swType = Type.GetTypeFromProgID("SldWorks.Application");
                if (swType == null)
                {
                    Log("ERROR: SolidWorks is not registered in COM. Exiting.");
                    return;
                }
                SldWorks swApp = (SldWorks)Activator.CreateInstance(swType);
                swApp.Visible = false;
                Log("✅ SolidWorks initialized successfully.");

                // Process SLD Files
                foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.slddrw"))
                {
                    Log($"Processing SLD file: {filePath}");
                    string pdfFilePath = ConvertSlddrwToPdf(swApp, filePath, outputDirectoryPath);
                    if (!string.IsNullOrEmpty(pdfFilePath))
                    {
                        ExtractTextFromPdf(pdfFilePath);
                    }
                }

                // Process TIFF Files
                foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.tif"))
                {
                    Log($"Processing TIFF file: {filePath}");
                    string pdfFilePath = ConvertTifToPdf(filePath, outputDirectoryPath);
                    if (!string.IsNullOrEmpty(pdfFilePath))
                    {
                        ExtractTextFromPdf(pdfFilePath);
                    }
                }

                // Process DWG Files
                foreach (string filePath in Directory.GetFiles(inputDirectoryPath, "*.dwg"))
                {
                    Log($"Processing DWG file: {filePath}");
                    string pdfFilePath = ConvertDwgToPdf(filePath, outputDirectoryPath);
                    if (!string.IsNullOrEmpty(pdfFilePath))
                    {
                        ExtractTextFromPdf(pdfFilePath);
                    }
                }

                swApp.ExitApp();
                swApp = null;
                GC.Collect();
                Log("✅ Program completed successfully.");
            }
            catch (Exception ex)
            {
                Log($"ERROR: Exception occurred in Main(): {ex.Message}");
            }
        }

        static string ConvertSlddrwToPdf(SldWorks swApp, string filePath, string outputDirectory)
        {
            try
            {
                Log($"Converting SLD file: {filePath}");

                if (!File.Exists(filePath))
                {
                    Log($"ERROR: File does not exist - {filePath}");
                    return null;
                }

                ModelDoc2 drawingDoc = (ModelDoc2)swApp.OpenDoc(filePath, (int)swDocumentTypes_e.swDocDRAWING);
                if (drawingDoc == null)
                {
                    Log($"ERROR: Failed to open SolidWorks drawing - {filePath}");
                    return null;
                }

                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
                int errors = 0, warnings = 0;

                bool saveResult = drawingDoc.Extension.SaveAs(pdfFilePath,
                                                             (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                                             (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                                             null,
                                                             ref errors,
                                                             ref warnings);

                swApp.CloseDoc(drawingDoc.GetTitle());

                if (saveResult && errors == 0)
                {
                    Log($"✅ Converted {filePath} to {pdfFilePath}");
                    return pdfFilePath;
                }
                else
                {
                    Log($"ERROR: Conversion failed for {filePath} - Errors: {errors}, Warnings: {warnings}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: Exception in ConvertSlddrwToPdf: {ex.Message}");
                return null;
            }
        }

        static string ConvertTifToPdf(string filePath, string outputDirectory)
        {
            try
            {
                Log($"Converting TIF file: {filePath}");

                if (!File.Exists(filePath))
                {
                    Log($"ERROR: File does not exist - {filePath}");
                    return null;
                }

                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                using (PdfSharpDocument document = new PdfSharpDocument())
                {
                    using (MagickImageCollection images = new MagickImageCollection(filePath))
                    {
                        foreach (MagickImage image in images)
                        {
                            PdfSharpPage page = document.AddPage();
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

                Log($"✅ Converted {filePath} to {pdfFilePath}");
                return pdfFilePath;
            }
            catch (Exception ex)
            {
                Log($"ERROR: Exception in ConvertTifToPdf: {ex.Message}");
                return null;
            }
        }

        static string ConvertDwgToPdf(string filePath, string outputDirectory)
        {
            try
            {
                Log($"Converting DWG file: {filePath}");

                AcadApplication acadApp = (AcadApplication)Activator.CreateInstance(Type.GetTypeFromProgID("AutoCAD.Application"), true);
                acadApp.Visible = false;

                AcadDocument acadDoc = acadApp.Documents.Open(filePath, false);
                System.Threading.Thread.Sleep(2000);

                string pdfFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(filePath) + ".pdf");
                acadDoc.Plot.PlotToFile(pdfFilePath, "DWG To PDF.pc3");

                acadDoc.Close(false);
                acadApp.Quit();

                Log($"✅ Converted {filePath} to {pdfFilePath}");
                return pdfFilePath;
            }
            catch (Exception ex)
            {
                Log($"ERROR: Exception in ConvertDwgToPdf: {ex.Message}");
                return null;
            }
        }
        static void ExtractTextFromPdf(string pdfFilePath)
        {
            try
            {
                Log($"Extracting text from {pdfFilePath}...");

                using (iTextPdfReader reader = new iTextPdfReader(pdfFilePath))
                using (iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(reader))
                {
                    string extractedText = "";
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        extractedText += iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i)) + "\n";
                    }

                    string txtFilePath = Path.ChangeExtension(pdfFilePath, ".txt");
                    File.WriteAllText(txtFilePath, extractedText);
                    Log($"✅ Extracted text from {pdfFilePath} to {txtFilePath}");
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR: Exception in ExtractTextFromPdf: {ex.Message}");
            }
        }
        static void Log(string message)
        {
            try
            {
                string logMessage = $"{DateTime.Now}: {message}";
                Console.WriteLine(logMessage);
                File.AppendAllText(logFile, logMessage + System.Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: Logging failed: {ex.Message}");
            }
        }

        static string SelectDirectory(string description)
        {
            using (FolderBrowserDialog folderBrowser = new FolderBrowserDialog())
            {
                folderBrowser.Description = description;
                folderBrowser.ShowNewFolderButton = false;
                return folderBrowser.ShowDialog() == DialogResult.OK ? folderBrowser.SelectedPath : null;
            }
        }
    }
}
