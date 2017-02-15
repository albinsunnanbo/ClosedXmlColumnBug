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
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Threading.Tasks;

    namespace OAS.Helpers.Excel
    {
        public static class ExcelExtensions
        {
            public static byte[] ToByteArray(this XLWorkbook workbook)
            {
                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    return memoryStream.ToArray();
                }
            }

            public static IXLWorksheet RemoveColumnFromTable(this IXLWorksheet workSheet, int columnNumber, string columnHeader)
            {
                var columnName = workSheet.Cell(4, columnNumber).GetString();
                if (columnName != columnHeader)
                {
                    throw new InvalidOperationException($"Tried to delete column {columnName}, expected '{columnHeader}'");
                }
                workSheet.Column(columnNumber).Delete(); // Delete empty the whole column

                return workSheet; // to enable fluent syntax
            }

            /// <summary>
            /// Format cells in worksheet to specific column types
            /// </summary>
            /// <param name="worksheet"></param>
            /// <param name="columnTypes"></param>
            /// <param name="beginningDataRow"></param>
            /// <param name="lastDataRow"></param>
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "beginningDataRow+1", Justification = "We won't go that far")]
            public static void FormatCellsInWorksheet(this IXLWorksheet worksheet, PropertyInfo[] columnTypes, int beginningDataRow, int lastDataRow)
            {
                worksheet.StyleTypes(columnTypes, beginningDataRow, lastDataRow);
                worksheet.Columns().AdjustToContents(beginningDataRow, lastDataRow);
                AddSpaceForDropDownMarker(worksheet);
            }

            public static void StyleTypes(this IXLWorksheet worksheet, PropertyInfo[] columnTypes, int beginningDataRow, int lastDataRow)
            {
                for (int i = 0; i < columnTypes.Count(); i++)
                {
                    if (columnTypes[i].PropertyType == typeof(TimeSpan) || columnTypes[i].PropertyType == typeof(TimeSpan?))
                    {
                        worksheet.Column(i + 1).Cells(beginningDataRow + 1, lastDataRow).Style.NumberFormat.NumberFormatId = 20; // H:mm
                    }
                    else if (columnTypes[i].PropertyType == typeof(bool))
                    {
                        worksheet.Column(i + 1).CellsUsed(ce => ce.Value as bool? == true).Value = "Ja";
                        worksheet.Column(i + 1).CellsUsed(ce => ce.Value as bool? == false).Value = "Nej";
                    }
                }
            }

            /// <summary>
            /// Adjust column with to accommodate for the drop down buttons
            /// </summary>
            /// <param name="worksheet"></param>
            public static void AddSpaceForDropDownMarker(this IXLWorksheet worksheet)
            {
                var lastCol = worksheet.LastColumnUsed().ColumnNumber();
                for (var colIdx = 1; colIdx <= lastCol; colIdx++)
                {
                    worksheet.Column(colIdx).Width += 2; // Add space for the drop down
                }
            }

        }
    }

}
