namespace GeneralSQLReporter.Helpers.Exporters
{
    using System;
    using System.IO;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Interaction Logic for the Helper Class <see cref="HtmlExporter"/>
    /// </summary>
    /// <remarks>
    /// Handles the Generating of the <see cref="ReportResultSet"/> into a HTML Document
    /// </remarks>
    internal static class HtmlExporter
    {
        /// <summary>
        /// <inheritdoc cref="SqlReportExporter.ExportHtml(ReportResultSet, string, bool, string)"/>
        /// </summary>
        internal static string ExportHtmlBase(ReportResultSet report,
            string templatePath = null,
            bool overwrite = true,
            string fileName = null)
        {
            SqlReportExporter.CheckExportFolder();

            var generatedHtml = HtmlExporter.GenerateBaseHtml(report, templatePath);
            var fullPath = string.IsNullOrWhiteSpace(fileName?.Trim()) ?
                Path.Combine(SqlReportExporter.OutputDirectory(), $"{Guid.NewGuid()}.html") :
                Path.Combine(SqlReportExporter.OutputDirectory(), fileName);

            if (File.Exists(fullPath) && overwrite)
            {
                File.Delete(fullPath);
            }

            File.WriteAllText(fullPath, generatedHtml);
            return fullPath;
        }

        /// <summary>
        /// Handles Generating the HTML Report contents.
        /// </summary>
        /// <param name="report">The <see cref="ReportResultSet"/> to generate the HTML Report of</param>
        /// <param name="templatePath">The path to the Template HTML to use if set</param>
        /// <returns>The generate HTML report content to be saved</returns>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the BaseHtml template resource was null or empty.</exception>
        internal static string GenerateBaseHtml(ReportResultSet report, string templatePath)
        {
            var usingTemplate = !string.IsNullOrWhiteSpace(templatePath?.Trim());

            var baseHtml = !usingTemplate ?
                Resource.html_template :
                File.ReadAllText(templatePath?.Trim());

            if (string.IsNullOrWhiteSpace(baseHtml))
            {
                throw new Exception($"Resource was missing, please raise an issue and inform me.");
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

            return baseHtml;
        }
    }
}