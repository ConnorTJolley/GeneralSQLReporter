namespace GeneralSQLReporter.Enums
{
    using System;

    /// <summary>
    /// Enumerator for the different available Report Output Formats
    /// </summary>
    public enum ReportFormat
    {
        /// <summary>
        /// Not Set, which would throw the <see cref="ArgumentException"/> when checking for a valid report
        /// </summary>
        NotSet = -1,

        /// <summary>
        /// Output Format in HTML in a generic plain HTML5 + CSS3 table with the column names and values
        /// </summary>
        Html = 0,

        /// <summary>
        /// Output Format in Excel in an .xlsx File
        /// </summary>
        ExcelXlsx = 1,

        /// <summary>
        /// Output Format in an .csv file
        /// </summary>
        Csv = 2,

        /// <summary>
        /// Output Format in an .pdf file
        /// </summary>
        Pdf = 3
    }
}