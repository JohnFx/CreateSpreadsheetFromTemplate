using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CreateSpreadsheetFromTemplate;
using System.IO;
using System.Diagnostics;


namespace CreateSpreadsheetFromTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            string rootPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, @"..\..\TestData\");
            string templateFilePath =  Path.Combine(rootPath,"TemplateTest.xlsx");
            string outputFilePath = Path.Combine(rootPath, "TemplateTest_Result.xlsx");
            try {
                if (File.Exists(outputFilePath)) { File.Delete(outputFilePath); }                
                FileInfo newFile = Excel.CreateFileWithTemplate(new FileInfo(templateFilePath), outputFilePath);
                Console.WriteLine(string.Format("Created file: {0}",outputFilePath));
                Process.Start(outputFilePath);
            }
        catch (Exception ex) {
            Console.WriteLine(ex.InnerException.ToString());
        }
    }
    }
}
