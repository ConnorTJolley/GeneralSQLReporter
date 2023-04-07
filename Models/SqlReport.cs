namespace GeneralSQLReporter.Models
{
    using System;
    using System.Collections.Generic;
    using GeneralSQLReporter.Helpers;
    using System.Net.Mail;
    using GeneralSQLReporter.Enums;

    public class SqlReport
    {
        /// <summary>
        /// Gets the SQL Query that is set against this <see cref="SqlReport"/> which is used by the 
        /// <see cref="SqlRepository"/> Helper class to run the Query and report the results.
        /// </summary>
        public string Query { get; internal set; }

        /// <summary>
        /// Gets the <see cref="ReportFormat"/> that is set against this <see cref="SqlReport"/>
        /// </summary>
        public ReportFormat OutputFormat { get; internal set; }

        /// <summary>
        /// Private List of Parsed <see cref="MailAddress"/>' from <see cref="EmailRecipients"/>
        /// </summary>
        private List<MailAddress> _emailAddresses = new List<MailAddress>();

        /// <summary>
        /// Backing field for the <see cref="EmailRecipients"/> Property
        /// </summary>
        private List<string> _emailRecipients = new List<string>();

        /// <summary>
        /// The List of Email Addresses that will recieve the Result of the Report via Email over SMTP.
        /// </summary>
        /// <remarks>
        /// This can be left empty if you do not wish to send the Report Result to the users Email.
        /// 
        /// If this is set, you must setup the <see cref="SmtpEmailSender"/> Class with the SMTP details
        /// </remarks>
        public List<string> EmailRecipients
        {
            get => this._emailRecipients;
            internal set
            {
                try
                {
                    if (value != null)
                    {
                        //// Clear list of Addresses first
                        this._emailAddresses = new List<MailAddress>();

                        this._emailRecipients = value;
                        value.ForEach(add =>
                        {
                            //// Iterate and parse each address to ensure it's valid
                            var address = new MailAddress(add);
                            this._emailAddresses.Add(address);
                        });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception occured when checking and updating EmailAddresses. " +
                        $"Please ensure they are all valid. Ex: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Gets or sets the value for Saving the Result of the SQL Report to the Local Machines Disk or not
        /// If this is set to <see cref="true"/> and the property <see cref="EmailRecipients"/> is set, 
        /// the locally created File will NOT be deleted after sending.
        /// If this is set to <see cref="false"/> the and the property <see cref="EmailRecipients"/> is empty 
        /// then an <see cref="ArgumentException"/> will be thrown when trying to Setup the <see cref="SqlRepository"/> with the report.
        /// </summary>
        public bool SaveToLocalDisk { get; internal set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="SqlReport"/> class for use by the 
        /// <see cref="SqlRepository"/> helper class to execute and report SQL Queries.
        /// </summary>
        /// <param name="query">The SQL Query to run</param>
        /// <param name="format">The desired Output <see cref="ReportFormat"/>, defaults to <see cref="ReportFormat.ExcelXls"/></param>
        /// <param name="saveToDisk"><see cref="true"/>/<see cref="false"/> for whether to save the output to Disk and Keep it, defaults to False</param>
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
        /// <param name="saveToDisk"><see cref="true"/>/<see cref="false"/> for whether to save the output to Disk and Keep it, defaults to False</param>
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