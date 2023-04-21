namespace GeneralSQLReporter.Helpers.Exporters
{
    using System;
    using System.IO;
    using System.Text;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Interaction Logic for the Helper Class <see cref="CsvExporter"/>
    /// </summary>
    /// <remarks>
    /// Handles the Generating of the <see cref="ReportResultSet"/> into a CSV Document
    /// </remarks>
    internal static class CsvExporter
    {
        /// <summary>
        /// <inheritdoc cref="SqlReportExporter.ExportCsv(ReportResultSet, bool, bool, string, char)"/>
        /// </summary>
        internal static string ExportCsv(ReportResultSet report,
            bool includeColumns,
            bool overwrite, 
            string fileName,
            char delimeter)
        {
            SqlReportExporter.CheckExportFolder();

            var fullPath = string.IsNullOrWhiteSpace(fileName?.Trim()) ?
                Path.Combine(SqlReportExporter.OutputDirectory(), $"{Guid.NewGuid()}.csv") :
                Path.Combine(SqlReportExporter.OutputDirectory(), fileName);

            var builder = new StringBuilder();

            if (includeColumns)
            {
                var colString = string.Empty;
                report.Columns.ForEach(col =>
                {
                    colString += $"{col.Name}{delimeter}"; //// Remove last delimeter
                });

                colString = colString.Remove(colString.Length - 1, 1);
                builder.Append(colString).Append(Environment.NewLine);
            }

            var rows = report.Rows;
            rows.ForEach(r =>
            {
                var fullRow = string.Empty;

                r.Values.ForEach(val =>
                {
                    fullRow += $"{val.Value}{delimeter}";
                });

                fullRow = fullRow.Remove(fullRow.Length - 1, 1); //// Remove last delimeter
                builder.Append(fullRow).Append(Environment.NewLine);
            });

            if (File.Exists(fullPath) && overwrite)
            {
                File.Delete(fullPath);
            }

            File.WriteAllText(fullPath, builder.ToString());
            return fullPath;
        }
    }
}