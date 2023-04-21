namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.IO;
    using System.Linq;
    using GeneralSQLReporter.Enums;
    using GeneralSQLReporter.Helpers.Exporters;
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
        /// Handles Retrieving the Path for the <see cref="SqlReportExporter._baseOutputDirectory"/> for where to Export Reports
        /// </summary>
        /// <returns>The Path of the Directory to Export Reports to</returns>
        public static string OutputDirectory() => SqlReportExporter._baseOutputDirectory;

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
        internal static void CheckExportFolder()
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
        /// <param name="overwrite">Whether to overwrite existing files or not, defaults to True</param>
        /// <param name="fileName">The Name to save the file as, e.g Report.html, if left empty will generate a GUID for the filename.</param>
        /// <returns>The Path of the Generated Report Document</returns>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the Template is empty</exception>
        public static string ExportHtml(ReportResultSet report, 
            string templatePath = null, 
            bool overwrite = true, 
            string fileName = null) =>
                HtmlExporter.ExportHtmlBase(report, templatePath, overwrite, fileName);

        /// <summary>
        /// Handles Export the <see cref="ReportResultSet"/> into a CSV Format
        /// </summary>
        /// <param name="report">The <see cref="ReportResultSet"/> to Export</param>
        /// <param name="includeColumns">Whether to Include the Columns in the CSV Export</param>
        /// <param name="overwrite">Wherther to overwrite existing files or not, default to True</param>
        /// <param name="fileName">The Name to save the file as, e.g Report.csv if left empty will generate  Guid for the filename.</param>
        /// <param name="delimeter">The Delimeter to use to separate the values, defaults to a comma.</param>
        /// <returns>The Path of the Generated Report Document</returns>
        public static string ExportCsv(ReportResultSet report,
            bool includeColumns = false,
            bool overwrite = true,
            string fileName = null,
            char delimeter = ',') =>
                CsvExporter.ExportCsv(report, includeColumns, overwrite, fileName, delimeter);

        private static string ExportPdf(ReportResultSet report, 
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