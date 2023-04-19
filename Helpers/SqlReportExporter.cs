namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.IO;
    using System.Linq;
    using GeneralSQLReporter.Enums;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Interaction Logic for the Helper class <see cref="SqlReportExporter"/>
    /// </summary>
    /// <remarks>
    /// Handles Generating the Various different <see cref="ReportFormat"/>s for the <see cref="GenericReport"/>
    /// </remarks>
    public static class SqlReportExporter
    {
        /// <summary>
        /// Private field for the Base Output Directory set by <see cref="SqlReportExporter.SetOutputDirectory(string)"/>
        /// </summary>
        private static string _baseOutputDirectory = $@"{AppContext.BaseDirectory}\GeneralSQLReporterOutputs\";

        /// <summary>
        /// Handles setting the Output Directory for all Reports to be Saved to
        /// </summary>
        /// <param name="outputDirectory">The Output Directory to Set, defaults to (<see cref="AppContext.BaseDirectory"/>\GeneralSQLReporterOutputs\) if empty.</param>
        /// <exception cref="ArgumentException">Throws an <see cref="ArgumentException"/> if the Path contains Invalid Characters</exception>
        public static void SetOutputDirectory(string outputDirectory)
        {
            var trimmed = outputDirectory.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                var defaultLocation = $@"{AppContext.BaseDirectory}\GeneralSQLReporterOutputs\";
                if (!Directory.Exists(defaultLocation))
                {
                    Directory.CreateDirectory(defaultLocation);
                }

                SqlReportExporter._baseOutputDirectory = defaultLocation;
            }

            var invalidChars = Path.GetInvalidPathChars();
            var invalidPath = false;
            invalidChars.ToList().ForEach(c =>
            {
                if (trimmed.Contains(c))
                {
                    invalidPath = true;
                    return;
                }
            });

            if (invalidPath)
            {
                throw new ArgumentException("Path contains Invalid Characters.");
            }

            try
            {
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                SqlReportExporter._baseOutputDirectory = outputDirectory;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Handles Checking and Ensuring the <see cref="SqlReportExporter._baseOutputDirectory"/> Directory Exists
        /// </summary>
        private static void CheckExportFolder()
        {
            if (!Directory.Exists(SqlReportExporter._baseOutputDirectory))
            {
                Directory.CreateDirectory(SqlReportExporter._baseOutputDirectory);
            }
        }

        /// <summary>
        /// Handles Exporting the <see cref="ReportResultSet"/> into a HTML format.
        /// </summary>
        /// <param name="report">The <see cref="ReportResultSet"/> to Export</param>
        /// <param name="templatePath">The Path to a matching Template HTML file, will default to included template.</param>
        /// <param name="overwrite">Whether to overwrite existing files or not, defaults to False</param>
        /// <returns>The Path of the Generated Report Document</returns>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Template is empty</exception>
        public static string ExportHtml(ReportResultSet report, 
            string templatePath = null, 
            bool overwrite = false)
        {
            SqlReportExporter.CheckExportFolder();

            var usingTemplate = !string.IsNullOrWhiteSpace(templatePath.Trim());

            var fileName = $"{Guid.NewGuid()}.html";
            var baseHtml = !usingTemplate ? 
                Resource.html_template : 
                File.ReadAllText(templatePath);

            if (string.IsNullOrWhiteSpace(baseHtml))
            {
                throw new Exception($"Resource was missing / empty, possibly corrupted .dll");
            }

            baseHtml = baseHtml.Replace("[COLCOUNT]", report.Columns.Count.ToString());
            baseHtml = baseHtml.Replace("[ROWCOUNT]", report.Rows.Count.ToString());

            var baseColumnVal = $"\t<td class=\"tg-hmp3\">[COL]</td>{Environment.NewLine}";
            var baseRowVal = $"\t<td class=\"tg-0lax\">[VAL]</td>{Environment.NewLine}";

            if (usingTemplate)
            {
                baseRowVal = baseRowVal.Replace(" class=\"tg-0lax\"", string.Empty);
                baseColumnVal = baseColumnVal.Replace(" class=\"tg-hmp3\"", string.Empty);
            }

            var fullColumnHtmlElement = string.Empty;
            var fullRowHtmlElement = string.Empty;
            report.Columns.ForEach(col =>
            {
                var replaced = baseColumnVal.Replace("[COL]", col.Name);
                fullColumnHtmlElement += replaced;
            });

            report.Rows.ForEach(row =>
            {
                var baseRow = $"<tr>[ROWVAL]</tr>";
                var fullRow = string.Empty;
                var values = row.Values;

                values.ForEach(val =>
                {
                    var replaced = baseRowVal.Replace("[VAL]", val.Value.ToString());
                    fullRow += replaced;
                });

                baseRow = baseRow.Replace("[ROWVAL]", fullRow);
                fullRowHtmlElement += baseRow;
            });

            baseHtml = baseHtml.Replace("[HEADERS]", fullColumnHtmlElement);
            baseHtml = baseHtml.Replace("[RECORDS]", fullRowHtmlElement);

            var fullPath = Path.Combine(SqlReportExporter._baseOutputDirectory, fileName);
            if (File.Exists(fullPath) && overwrite)
            {
                File.Delete(fullPath);
            }

            File.WriteAllText(fullPath, baseHtml);
            return fullPath;
        }

        private static string ExportPdf(ReportResultSet report, 
            string templatePath = null, 
            bool overwrite = false)
        {
            SqlReportExporter.CheckExportFolder();

            throw new NotImplementedException();
        }

        private static string ExportCsv(ReportResultSet report, 
            string templatePath = null, 
            bool overwrite = false)
        {
            SqlReportExporter.CheckExportFolder();

            throw new NotImplementedException();
        }

        private static string ExportExcel(ReportResultSet report, 
            ReportFormat format, 
            string templatePath = null, 
            bool overwrite = false)
        {
            SqlReportExporter.CheckExportFolder();

            throw new NotImplementedException();
        }
    }
}