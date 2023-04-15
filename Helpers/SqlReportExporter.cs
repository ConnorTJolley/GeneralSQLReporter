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
        private static string _baseOutputDirectory;

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
    }
}