using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace CreateSpreadsheetFromTemplate
{
    public static class Excel
    {
        public static FileInfo CreateFileWithTemplate(FileInfo templateFile, string newFilePath)
        {
            try
            {
                string ErrorMessage = string.Empty;
                ExcelPackage xlPackage;
                FileInfo newFile = new FileInfo(newFilePath);

                using (xlPackage = new ExcelPackage(newFile, templateFile))
                {

                    ///* Set metadata */
                    xlPackage.Workbook.Properties.Title = "Sample XLS file from template";
                    xlPackage.Workbook.Properties.Author = "John Fuex";                                     
                    xlPackage.Workbook.Properties.Company = "Deloitte Discovery";

                    foreach (ExcelWorksheet ws in xlPackage.Workbook.Worksheets) {
                        for (int row = 9; row<=20; row++) { 
                            ws.Cells["A" + row].Value = "John";
                            ws.Cells["B" + row].Value = "Fuex";
                            ws.Cells["C" + row].Formula = "=CONCATENATE(A9,\" \", B9)";
                            ws.Cells["D" + row].Value = 100000 * row;
                            ws.Cells["D" + row].Style.Numberformat.Format = "$#,##0.00";                                   
                        }
                        ws.Cells["D9:D20"].AutoFitColumns();
                        ws.Cells["D8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }
                    xlPackage.Workbook.View.ActiveTab = 0;                    
                    xlPackage.Save();                    
                }
                return newFile;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                return null;
            }
        }
    }
}
