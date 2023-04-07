namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Static Helper for Creating and Sending Emails via SMTP to Recipients on <see cref="SqlReport.EmailRecipients"/>
    /// </summary>
    /// <remarks>
    /// If the <see cref="SqlReport.EmailRecipients"/> is not empty but the <see cref="SmtpEmailSender"/> has NOT
    /// been configured, an <see cref="ArgumentException"/> will be thrown for an Invalid <see cref="SqlReport"/>
    /// </remarks>
    public static class SmtpEmailSender
    {
        /// <summary>
        /// Private instance of the <see cref="SmtpClient"/>.
        /// </summary>
        private static SmtpClient _smtpClient;
        private static string _fromAddress;

        /// <summary>
        /// Initializes the static instance of the <see cref="SmtpEmailSender"/> with an empty <see cref="SmtpClient"/>
        /// </summary>
        static SmtpEmailSender()
        {
            SmtpEmailSender._smtpClient = new SmtpClient();
        }

        public static bool Setup(string host, int port, NetworkCredential credentials, string fromAddress)
        {
            SmtpEmailSender._smtpClient.Host = host;
            SmtpEmailSender._smtpClient.Port = port;
            SmtpEmailSender._smtpClient.UseDefaultCredentials = false;
            SmtpEmailSender._fromAddress = fromAddress.Trim();

            SmtpEmailSender._smtpClient.Credentials = credentials;

            return SmtpEmailSender.IsSetup();
        }

        /// <summary>
        /// Checks to see whether or not the <see cref="SmtpEmailSender"/> has been configured correctly
        /// </summary>
        /// <returns><see cref="true"/> if setup. <see cref="false"/> if not</returns>
        /// <remarks>
        /// This is used when checking for a valid <see cref="SqlReport"/> 
        /// if <see cref="SqlReport.EmailRecipients"/> is NOT empty before running query.
        /// 
        /// It is advised that this is checked before running any 
        /// report that has any <see cref="SqlReport.EmailRecipients"/> 
        /// </remarks>
        public static bool IsSetup()
        {
            if (SmtpEmailSender._smtpClient == null)
            {
                return false;
            }

            var host = SmtpEmailSender._smtpClient.Host;
            var port = SmtpEmailSender._smtpClient.Port;
            
            if (string.IsNullOrWhiteSpace(host.ToString()) && port == 25)
            {
                //// No settings have been changed
                return false;
            }

            var creds = SmtpEmailSender._smtpClient.Credentials;
            if (creds == null)
            {
                return false;
            }

            try
            {
                var _ = new MailAddress(SmtpEmailSender._fromAddress);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception checking FromAddress is Valid. Ex: {ex.Message}");
                return false;
            }
        }

        public static bool SendReportResultsToEmails(SqlReport report, ReportResultSet result)
        {
            if (!report.EmailRecipients.Any())
            {
                throw new ArgumentException("Email Recipients was empty.");
            }

            var successCount = 0;

            try
            {
                ////TODO Get the report output

                var subject = $"Report Results - {DateTime.Now.ToShortDateString()}";
                var body = $"A report has been generated and attached to this email for your viewing.";

                report.EmailRecipients.ForEach(recipient =>
                {
                    var mailMessage = new MailMessage(SmtpEmailSender._fromAddress, 
                        recipient, 
                        subject,
                        body);

                    ////TODO add email attachment

                    SmtpEmailSender._smtpClient.Send(mailMessage);
                    successCount++;
                });

                return successCount == report.EmailRecipients.Count;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}