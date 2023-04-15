namespace GeneralSQLReporter.Models
{
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.Net.Mail;
    using GeneralSQLReporter.Enums;
    using GeneralSQLReporter.Helpers;

    /// <summary>
    /// Interaction Logic for the <see cref="GenericReport"/> Class.
    /// </summary>
    /// <remarks>
    /// Contains Shared Properties between both the <see cref="SqlReport"/> and the <see cref="StoredProcedureReport"/> classes.
    /// </remarks>
    public class GenericReport
    {
        /// <summary>
        /// Gets the <see cref="ReportFormat"/> that is set against this <see cref="GenericReport"/>
        /// </summary>0
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
        /// Gets the <see cref="List{SqlParameter}"/>'s for the Stored Procedure Report
        /// </summary>
        public List<SqlParameter> Parameters { get; internal set; }

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
        /// If this is set to True and the property <see cref="EmailRecipients"/> is set, 
        /// the locally created File will NOT be deleted after sending.
        /// If this is set to False the and the property <see cref="EmailRecipients"/> is empty 
        /// then an <see cref="ArgumentException"/> will be thrown when trying to Setup the <see cref="SqlRepository"/> with the report.
        /// </summary>
        public bool SaveToLocalDisk { get; internal set; }
    }
}