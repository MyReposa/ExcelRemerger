using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Reflection;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ExcelRemerger
{
    class Demerger
    {
        Logger logger;

        Microsoft.Office.Interop.Excel.Application excelApp;
        Workbook macroWorkbook;
        string filePath;
        string fileName;
        string macroFileName;
        string workDirectory;
        List<WorksheetData> allWorksheetData = new List<WorksheetData>();

        public Demerger(Logger logger)
        {
            this.logger = logger;
        }

        public void Work(string filePath)
        {
            try
            {
                this.filePath = filePath;
                this.workDirectory = Path.GetDirectoryName(filePath);
                this.fileName = Path.GetFileName(filePath);
                this.macroFileName = Path.GetFileNameWithoutExtension(this.fileName) + "_RemergerMacro.xlsm";
                if (GetMergedCells())
                {
                    CreateMacroFile();
                    CreateActualMacro();
                    SaveAndCloseMacroFile();
                }

                System.Diagnostics.Process.Start($@"{workDirectory}\{this.macroFileName}");
            }
            catch (Exception error)
            {
                logger.Log($"Error:\r\n" +
                           $"{error.Message}");
            }
        }

        bool GetMergedCells()
        {
            try
            {
                logger.Log("Reading data of merged cells...");
                if (Directory.Exists($@"{ workDirectory}\_rmrgtemp") == false)
                {
                    Directory.CreateDirectory($@"{ workDirectory}\_rmrgtemp");
                }

                using (ZipArchive archive = ZipFile.OpenRead(filePath))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName == "xl/workbook.xml")
                        {
                            entry.ExtractToFile($@"{workDirectory}\_rmrgtemp\{entry.Name}", true);

                            XDocument xmlFile = XDocument.Load($@"{workDirectory}\_rmrgtemp\{entry.Name}");
                            XNamespace defaultXmlNamespace = xmlFile.Root.GetDefaultNamespace();

                            foreach (XElement sheet in xmlFile.Root.Elements(defaultXmlNamespace + "sheets").Elements(defaultXmlNamespace + "sheet"))
                            {
                                WorksheetData worksheetDataToAdd = new WorksheetData();
                                worksheetDataToAdd.name = sheet.Attribute("name").Value;
                                this.allWorksheetData.Add(worksheetDataToAdd);
                            }
                        }
                    }

                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.StartsWith(@"xl/worksheets/sheet") && entry.Name.EndsWith(".xml"))
                        {
                            entry.ExtractToFile($@"{workDirectory}\_rmrgtemp\{entry.Name}", true);
                            XDocument xmlFile = XDocument.Load($@"{workDirectory}\_rmrgtemp\{entry.Name}");
                            XNamespace defaultXmlNamespace = xmlFile.Root.GetDefaultNamespace();
                            int sheetNumber = Int32.Parse(Regex.Replace(entry.Name, $@"sheet|\.xml", ""));

                            foreach (XElement area in xmlFile.Root.Elements(defaultXmlNamespace + "mergeCells").Elements(defaultXmlNamespace + "mergeCell"))
                            {
                                allWorksheetData[sheetNumber - 1].mergedAreas.Add(area.Attribute("ref").Value);
                            }
                        }
                    }
                }
                Directory.Delete($@"{ workDirectory}\_rmrgtemp", true);
                logger.Log("Data reading completed.");
                return true;
            }
            catch (Exception error)
            {
                logger.Log($"Error:\r\n" +
                           $"{error.Message}");
                return false;
            }
        }

        void CreateMacroFile()
        {
            logger.Log("Creating macro file...");
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            this.macroWorkbook = this.excelApp.Workbooks.Add();
            Worksheet macroInfoWorksheet = macroWorkbook.Worksheets.Item[1];
            macroInfoWorksheet.Columns[1].ColumnWidth = 250;
            macroInfoWorksheet.Rows[1].RowHeight = 50;
            macroInfoWorksheet.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            macroInfoWorksheet.Cells[1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            macroInfoWorksheet.Cells[1, 1].Font.Size = 14;
            macroInfoWorksheet.Cells[1, 1].Font.Bold = true;
            macroInfoWorksheet.Cells[1, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
            macroInfoWorksheet.Cells[1, 1] = $"Remerger macro file for merging back previously unmerged cells\r\nCreated for file \"{fileName}\"";
            macroInfoWorksheet.Rows[2].RowHeight = 30;
            macroInfoWorksheet.Cells[2, 1].Font.Bold = true;
            macroInfoWorksheet.Cells[2, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            macroInfoWorksheet.Cells[2, 1] = "How to use:";
            macroInfoWorksheet.Cells[4, 1] = "Prep:";
            macroInfoWorksheet.Cells[5, 1] = "1. Use Remerger tool on Excel file containing merged cells to generate Remerger macro file. (Note: You just did that. This is Remerger macro file.)";
            macroInfoWorksheet.Cells[6, 1] = $"2. Unmerge any cells you need in \"{fileName}\" and prepare the file to translation in standard way.";
            macroInfoWorksheet.Cells[8, 1] = "Post\\Dummy:";
            macroInfoWorksheet.Cells[9, 1] = "1. Keep this file open anywhere in the background. Make sure permission for macro usage has been granted.";
            macroInfoWorksheet.Cells[10, 1] = $"2. Open file you wish to merge back. Keep in mind that this particular macro was generated for \"{fileName}\".";
            macroInfoWorksheet.Cells[11, 1] = "3. Make sure file to merge is active. Selection doesn't matter.";
            macroInfoWorksheet.Cells[12, 1] = "4. Launch merging macro by CTRL+R hotkey or by manually selecting from macros library.";
            macroInfoWorksheet.Cells[13, 1] = "5. That should merge cells back to original form.";
            logger.Log("Macro file created.");
        }

        void CreateActualMacro()
        {
            logger.Log("Writing macro...");
            _VBComponent vbComponent = macroWorkbook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
            string sCode =
                $"Sub Remerge()\r\n" +
                $"'\r\n" +
                $"' Remerger macro created for file \"{this.fileName}\"\r\n" +
                $"'\r\n" +
                $"' Keyboard Shortcut: Ctrl+r\r\n" +
                $"'\r\n" +
                $"\tDim StartingSheet As Worksheet\r\n" +
                $"\tDim SelRange As Range\r\n" +
                $"\tSet StartingSheet = ActiveSheet\r\n\r\n";

            foreach (WorksheetData sheetData in allWorksheetData)
            {
                sCode += $"\t\tSheets(\"{sheetData.name}\").Select\r\n" +
                         $"\t\t\tSet SelRange = Selection\r\n";

                foreach (string area in sheetData.mergedAreas)
                {
                    sCode += $"\t\t\tRange(\"{area}\").Select\r\n" +
                             $"\t\t\t\tSelection.Merge\r\n" +
                             $"\t\t\t\tSelection.HorizontalAlignment = xlCenter\r\n";
                }

                sCode += $"\t\t\tSelRange.Select\r\n\r\n";
            }
            sCode += $"\t\tStartingSheet.Select\r\n" +
                     $"End Sub";
            vbComponent.CodeModule.AddFromString(sCode);
            excelApp.MacroOptions("Remerge", $"Remerger macro for merging back unmerged cells. Created for file \"{this.fileName}\".", Missing.Value, Missing.Value, true, "r", Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            logger.Log("Macro writing completed.");
        }

        void SaveAndCloseMacroFile()
        {
            logger.Log("Saving macro file...");
            try
            {
                macroWorkbook.SaveAs($@"{this.workDirectory}\{this.macroFileName}", XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                this.macroWorkbook.Close();
                this.excelApp.Quit();

                //ReleaseObject(worksheet);
                ReleaseObject(macroWorkbook);
                ReleaseObject(excelApp);

                logger.Log("Macro file saved:\r\n" +
                           $@"{workDirectory}\{this.macroFileName}");
            }
            catch (Exception error)
            {
                logger.Log($"Error: {error.Message}");
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
