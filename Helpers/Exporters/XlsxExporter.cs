namespace GeneralSQLReporter.Helpers.Exporters
{
    using System;
    using System.IO;
    using Path = System.IO.Path;
    using GeneralSQLReporter.Models;
    using Spire.Xls;

    /// <summary>
    /// Interaction logic for the <see cref="XlsxExporter"/> Class.
    /// </summary>
    /// <remarks>
    /// Used by the <see cref="SqlReportExporter"/> to generate Excel (xlsx) documents based on the <see cref="ReportResultSet"/>'s
    /// </remarks>
    public static class XlsxExporter
    {
        /// <summary>
        /// <inheritdoc cref="SqlReportExporter.ExportExcel(ReportResultSet, string, bool, string, int, int, int, int, bool)"/>
        /// </summary>
        internal static string ExportXlsx(ReportResultSet report, 
            string templatePath = null,
            bool overwrite = true,
            string fileName = null,
            int sheetNumber = 0,
            int headerRow = 1,
            int firstRow = 2,
            int firstCol = 1,
            bool autofitCols = true)
        {
            SqlReportExporter.CheckExportFolder();

            var workbook = XlsxExporter.GenerateWorkbook(report, 
                templatePath, 
                sheetNumber, 
                headerRow, 
                firstRow, 
                firstCol, 
                autofitCols);

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

        /// <summary>
        /// Generalised Exporting of the <see cref="ReportResultSet"/> into an Excel Format
        /// </summary>
        /// <param name="report">The <see cref="ReportResultSet"/> to export</param>
        /// <param name="templatePath">The FilePath to a Template Excel file to use if applicable, if left empty will default to the Included Template</param>
        /// <param name="sheetNumber">The Index for the Sheet to use, defaults to 0 (first sheet) if left empty</param>
        /// <param name="headerRow">The Index for the Header Row, defaults to 1 if left empty</param>
        /// <param name="firstRow">The Index of the First Row for the record, defaults to 2 if left empty</param>
        /// <param name="firstCol">The Index of the FIrst Column for the headers/cell values, defaults to 1 if left empty</param>
        /// <param name="autofitCols">Value to indicate whether to autofit columns or not for the report.</param>
        /// <returns>The File Path for the Generated Excel Document</returns>
        internal static Workbook GenerateWorkbook(ReportResultSet report, 
            string templatePath = null,
            int sheetNumber = 0,
            int headerRow = 1,
            int firstRow = 2,
            int firstCol = 1,
            bool autofitCols = true)
        {
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

            //// Get range for the first row, to copy and styling set against it
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

            return workbook;
        }
    }
}