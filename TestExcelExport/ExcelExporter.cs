using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelExport
{
    using ClosedXML.Excel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;

    namespace OAS.Helpers.Excel
    {
        public static class ExcelExporter
        {
            /// <summary>
            /// The content type to use when returning xlsx files
            /// </summary>
            public const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            public static XLWorkbook CreateWorkbookFromTypedRows<T>(IEnumerable<T> rows, string sheetName, string header, string subHeader)
            {
                var columnTypes = typeof(T).GetProperties();

                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(sheetName);

                var headerCell = worksheet.Cell(1, 1);
                headerCell.Value = header;
                headerCell.Style.Font.SetBold().Font.SetFontSize(18);

                var subHeaderCell = worksheet.Cell(2, 1);
                subHeaderCell.Value = subHeader;

                var printDateCell = worksheet.Cell(3, 1);
                printDateCell.Value = "Exportdate: " + DateTime.Now;

                var startRow = worksheet.LastRowUsed().RowNumber() + 1;
                if (rows.Any())
                {
                    worksheet.Cell(startRow, 1).InsertTable(rows);
                }
                else
                {
                    var cell = worksheet.Cell(startRow + 1, 1);
                    cell.Value = "Nothing to export";
                    cell.Style.Font.Italic = true;
                }
                var endRow = worksheet.LastRowUsed().RowNumber();

                worksheet.FormatCellsInWorksheet(columnTypes, startRow, endRow);

                return workbook;
            }
        }
    }

}
