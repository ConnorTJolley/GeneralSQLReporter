namespace GeneralSQLReporter.Models
{
    using System.Collections.Generic;
    using GeneralSQLReporter.Helpers;
    using GeneralSQLReporter.Enums;

    /// <summary>
    /// Interaction Logic for the <see cref="SqlReport"/> Class.
    /// </summary>
    /// <remarks>
    /// Contains Properties and various constructors for the <see cref="SqlReport"/> class which is used when calling
    /// <see cref="SqlRepository.RunSingleReport(GenericReport)"/> or <see cref="SqlRepository.RunSingleReportAsync(GenericReport)"/>
    /// </remarks>
    public class SqlReport : GenericReport
    {
        /// <summary>
        /// Gets the SQL Query that is set against this <see cref="SqlReport"/> which is used by the 
        /// <see cref="SqlRepository"/> Helper class to run the Query and report the results.
        /// </summary>
        public string Query { get; internal set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlReport"/> class for use by the 
        /// <see cref="SqlRepository"/> helper class to execute and report SQL Queries.
        /// </summary>
        /// <param name="query">The SQL Query to run</param>
        /// <param name="format">The desired Output <see cref="ReportFormat"/>, defaults to <see cref="ReportFormat.ExcelXls"/></param>
        /// <param name="saveToDisk">True/False for whether to save the output to Disk and Keep it, defaults to False</param>
        /// <remarks>
        /// Bare-bones Initializer for the <see cref="SqlReport"/>, this is best used 
        /// when you are wanting to run a report quickly and save to disk or keep in memory.
        /// 
        /// E.g a user clicks a Report button within an App / Website, run the query, store the Report or not and 
        /// process the output however required.
        /// </remarks>
        public SqlReport(string query, ReportFormat format = ReportFormat.ExcelXls, bool saveToDisk = false)
        {
            this.Query = query;
            this.OutputFormat = format;
            this.SaveToLocalDisk = saveToDisk;

            this.EmailRecipients = new List<string>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlReport"/> class for use by the 
        /// <see cref="SqlRepository"/> helper class to execute and report SQL Queries.
        /// </summary>
        /// <param name="query">The SQL Query to run</param>
        /// <param name="format">The desired Output <see cref="ReportFormat"/>, defaults to <see cref="ReportFormat.ExcelXls"/></param>
        /// <param name="emailRecipients">
        /// The List of Email Addresses to send the Report to, 
        /// if populated <see cref="SmtpEmailSender"/> needs to be Configured.
        /// </param>
        /// <param name="saveToDisk">True/False for whether to save the output to Disk and Keep it, defaults to False</param>
        public SqlReport(string query, 
            ReportFormat format, 
            List<string> emailRecipients, 
            bool saveToDisk = false)
        {
            this.Query = query;
            this.OutputFormat = format;
            this.SaveToLocalDisk = saveToDisk;
            this.EmailRecipients = emailRecipients;
        }
    }
}