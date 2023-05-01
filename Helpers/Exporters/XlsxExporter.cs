using GeneralSQLReporter.Models;
using Spire.Pdf.Exporting.XPS.Schema;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Path = System.IO.Path;

namespace GeneralSQLReporter.Helpers.Exporters
{
    public static class XlsxExporter
    {
        internal static string ExportXlsx(ReportResultSet report, 
            string templatePath,
            bool overwrite,
            string fileName = null,
            int sheetNumber = 0,
            int headerRow = 1,
            int firstRow = 2,
            int firstCol = 1,
            bool autofitCols = true)
        {
            SqlReportExporter.CheckExportFolder();

            var usingTemplate = !string.IsNullOrWhiteSpace(templatePath?.Trim());
            var workbook = new Workbook();
            if (usingTemplate)
            {
                workbook.LoadFromFile(templatePath);
            }
            else
            {
                var stream = new MemoryStream();
                var templateBytes = Resource.excel_template;
                stream.Write(templateBytes, 0, templateBytes.Length);
                stream.Seek(0, SeekOrigin.Begin);
                workbook.LoadFromStream(stream);
                stream.Close();
                stream.Dispose();
            }

            //// Set active sheet
            workbook.ActiveSheetIndex = sheetNumber;
            var sheet = workbook.ActiveSheet;

            var range = sheet.Range[headerRow, firstCol, firstRow, firstCol];
            var colCount = report.Columns.Count - 1;

            var cols = report.Columns;
            var i = firstCol;
            var colIndex = firstCol;
            cols.ForEach(c =>
            {
                sheet.CopyColumn(range, sheet, colIndex, CopyRangeOptions.CopyStyles);
                sheet.SetCellValue(headerRow, colIndex, c.Name);
                colIndex++;
            });

            var endCol = colIndex;

            //// Copy row styling
            var rowRange = sheet.Range[firstRow, firstCol, firstRow, endCol];

            var rows = report.Rows;
            var r = firstRow;
            rows.ForEach(rec =>
            {
                //// Copy row formatting first
                sheet.CopyRow(rowRange, sheet, r, CopyRangeOptions.CopyStyles);

                colIndex = firstCol;

                rec.Values.ForEach(val =>
                {
                    sheet.SetCellValue(r, colIndex, val.Value.ToString());
                    colIndex++;
                });

                r++;
            });

            if (autofitCols)
            {
                var col = firstCol;
                while (firstCol != endCol)
                {
                    sheet.AutoFitColumn(col);
                    firstCol++;
                }
            }
            
            var fullPath = string.IsNullOrWhiteSpace(fileName?.Trim()) ?
                Path.Combine(SqlReportExporter.OutputDirectory(), $"{Guid.NewGuid()}.xlsx") :
                Path.Combine(SqlReportExporter.OutputDirectory(), fileName);

            if (File.Exists(fullPath) && overwrite)
            {
                File.Delete(fullPath);
            }

            workbook.SaveToFile(fullPath);

            return fullPath;
        }
    }
}