using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestExcelExport.OAS.Helpers.Excel;

namespace TestExcelExport
{
    class Program
    {
        static void Main(string[] args)
        {
            var doesntWorkModel = new List<DoesntWorkModel>
            {
                new DoesntWorkModel
                {
                     String1 = "Hello",
                     String2 = "World",
                }
            };

            using (var report = ExcelExporter.CreateWorkbookFromTypedRows(doesntWorkModel, "Test", "Test", $"Exported"))
            {
                // BEGIN Hack to remove unwanted columns
                // Could possibly be solved more generic in the future with an ExcelIgnore attribute or something like that
                report.Worksheet(1)
                    .RemoveColumnFromTable(2, "The Int");
                // END Hack
                report.SaveAs("DoesntWork.xlsx");
            }

            var doWorkModel = new List<DoWorkModel>
            {
                new DoWorkModel
                {
                     String1 = "Hello",
                     String2 = "World",
                }
            };

            using (var report = ExcelExporter.CreateWorkbookFromTypedRows(doWorkModel, "Test", "Test", $"Exported"))
            {
                // BEGIN Hack to remove unwanted columns
                // Could possibly be solved more generic in the future with an ExcelIgnore attribute or something like that
                report.Worksheet(1)
                    .RemoveColumnFromTable(3, "The Int");
                // END Hack
                report.SaveAs("DoWork.xlsx");
            }
        }
    }
}
