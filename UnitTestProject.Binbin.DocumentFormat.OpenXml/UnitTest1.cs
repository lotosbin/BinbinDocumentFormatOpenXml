using System;
using System.Diagnostics;
using Binbin.DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject.Binbin.DocumentFormat.OpenXml
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        [DeploymentItem("ExcelTemplates/Book1.xlsx", "Exports")]
        [DeploymentItem("ExcelTemplates/Book2.xlsx", "Exports")]
        public void TestMethod1()
        {

            const int count = 1000;
            {
                //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "./Exports/Book1.xlsx");
                const string filePath = "./Exports/Book1.xlsx";
                Debug.WriteLine(filePath);
                var start = DateTime.Now;
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
                {
                    for (int i = 0; i < count; i++)
                    {
                        document.UpdateCell( "Sheet1", (uint)(i + 1), "A",CellValues.String, "test");
                    }
                }
                var end = DateTime.Now;
                Debug.WriteLine((end - start));
            }
            {
                //var filePath = Path.Combine(Directory.GetCurrentDirectory(), "./Exports/Book2.xlsx");
                const string filePath = "./Exports/Book2.xlsx";
                Debug.WriteLine(filePath);
                var start = DateTime.Now;
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
                {
                    
                    var worksheet = SpreadsheetDocumentExtension.GetWorksheet(document, "Sheet1");
                    for (int i = 0; i < count; i++)
                    {
                        worksheet.UpdateCell((uint)(i + 1), "A", CellValues.String, "test");
                    }
                    document.Dispose();
                    //worksheet.Save();
                }
                var end = DateTime.Now;
                Debug.WriteLine((end - start));
            }

        }
    }
}
