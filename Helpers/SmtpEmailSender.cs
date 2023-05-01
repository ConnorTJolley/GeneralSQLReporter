namespace GeneralSQLReporter.Helpers
{
    using System;
    using System.Linq;
    using System.Net;
    using System.Net.Mail;
    using System.Threading.Tasks;
    using GeneralSQLReporter.Enums;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Static Helper for Creating and Sending Emails via SMTP to Recipients on <see cref="GenericReport.EmailRecipients"/>
    /// </summary>
    /// <remarks>
    /// If the <see cref="GenericReport.EmailRecipients"/> is not empty but the <see cref="SmtpEmailSender"/> has NOT
    /// been configured, an <see cref="ArgumentException"/> will be thrown for an Invalid <see cref="SqlReport"/>
    /// </remarks>
    public static class SmtpEmailSender
    {
        /// <summary>
        /// Private instance of the <see cref="SmtpClient"/>.
        /// </summary>
        private static SmtpClient _smtpClient;

        /// <summary>
        /// Private storage field for the From Email Address
        /// </summary>
        private static string _fromAddress;

        /// <summary>
        /// Handles the Setting Up of the <see cref="SmtpEmailSender"/> and validating that the backing <see cref="SmtpClient"/> is setup
        /// </summary>
        /// <param name="host">The SMTP host address, e.g smtp.gmail.com</param>
        /// <param name="credentials">The <see cref="NetworkCredential"/> for the SMTP Authentication</param>
        /// <param name="fromAddress">The Email Address to send the Emails As, e.g noreplay@address.com</param>
        /// <param name="port">The SMTP port, e.g 587 (defaults to 465)</param>
        /// <param name="requiresSsl">Value indicating if the SMTP requires SSL or not, defaults to True</param>
        /// <returns>True if Setup, False if not</returns>
        public static bool Setup(string host, 
            NetworkCredential credentials,
            string fromAddress,
            int port = 465, 
            bool requiresSsl = true)
        {
            SmtpEmailSender._smtpClient = new SmtpClient
            {
                Host = host,
                Port = port,
                EnableSsl = requiresSsl,
                UseDefaultCredentials = false
            };

            SmtpEmailSender._fromAddress = fromAddress;
            SmtpEmailSender._smtpClient.Credentials = credentials;

            return SmtpEmailSender.IsSetup();
        }

        /// <summary>
        /// Checks to see whether or not the <see cref="SmtpEmailSender"/> has been configured correctly
        /// </summary>
        /// <returns>True if setup. False if not</returns>
        /// <remarks>
        /// This is used when checking for a valid <see cref="SqlReport"/> 
        /// if <see cref="GenericReport.EmailRecipients"/> is NOT empty before running query.
        /// 
        /// It is advised that this is checked before running any 
        /// report that has any <see cref="GenericReport.EmailRecipients"/> 
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

        /// <summary>
        /// Handles Sending the <see cref="ReportResultSet"/> to the Recipient(s) Asynchronously
        /// </summary>
        /// <param name="result">The resulting <see cref="ReportResultSet"/></param>
        /// <param name="subject">The Subject for the Email</param>
        /// <param name="body">The Body for the Email</param>
        /// <param name="isBodyHtml">Value to indicate whether the Body is HTML or not, defaults to True</param>
        /// <param name="filePath">The Path to the Generated Report, if left Empty, it will generate the report itself</param>
        /// <returns>True if successfully sent to ALL recipients, False if not</returns>
        /// <exception cref="ArgumentException">
        /// Throws an <see cref="ArgumentNullException"/> if there are 0 
        /// Email Recipients Setup on the <see cref="GenericReport"/>
        /// </exception>
        public static async Task<bool> SendReportResultsToEmailsAsync(ReportResultSet result,
            string subject,
            string body,
            bool isBodyHtml = true,
            string filePath = null)
        {
            var report = result.ReportUsed;
            if (!report.EmailRecipients.Any())
            {
                throw new ArgumentException("Email Recipients was empty.");
            }

            var successCount = 0;

            try
            {
                filePath = SmtpEmailSender.HandleFilePath(filePath, result);
                var fileName = filePath.Substring(filePath.LastIndexOf('\\') + 1);

                var attachment = new Attachment(filePath)
                {
                    Name = fileName
                };

                foreach (var recipient in  report.EmailRecipients)
                {
                    var mailMessage = new MailMessage(SmtpEmailSender._fromAddress,
                        recipient,
                        subject,
                        body)
                    {
                        IsBodyHtml = isBodyHtml
                    };

                    //// Add attachment for the report
                    mailMessage.Attachments.Add(attachment);

                    await SmtpEmailSender._smtpClient.SendMailAsync(mailMessage);
                    successCount++;
                }

                return successCount == report.EmailRecipients.Count;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Handles Ensuring there is an Attachment to Report
        /// </summary>
        /// <param name="filePath">The path to the Pre-Existing file which can be left empty / null</param>
        /// <param name="result">The <see cref="ReportResultSet"/> which will be used to generate the output 
        /// if <paramref name="filePath"/> is left empty.</param>
        /// <returns>The Path to the generated output File</returns>
        /// <exception cref="Exception">Throws an <see cref="Exception"/> if the <see cref="ReportFormat"/> is 
        /// set to <see cref="ReportFormat.NotSet"/></exception>
        private static string HandleFilePath(string filePath, ReportResultSet result)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                //// No pre-generated attachment, need generate and use
                var type = result.ReportUsed.OutputFormat;
                switch (type)
                {
                    case ReportFormat.Html:
                        filePath = SqlReportExporter.ExportHtml(result);
                        break;

                    case ReportFormat.Csv:
                        filePath = SqlReportExporter.ExportCsv(result);
                        break;

                    case ReportFormat.ExcelXlsx:
                        //filePath = SqlReportExporter.ExportHtml(result);
                        break;

                    case ReportFormat.Pdf:
                        //filePath = SqlReportExporter.ExportHtml(result);
                        break;

                    case ReportFormat.NotSet:
                        throw new Exception("Report Output Format was 'NotSet'");
                }
            }

            return filePath;
        }

        /// <summary>
        /// Handles Sending the <see cref="ReportResultSet"/> to the Recipient(s) Synchronously
        /// </summary>
        /// <param name="result">The resulting <see cref="ReportResultSet"/></param>
        /// <param name="subject">The Subject for the Email</param>
        /// <param name="body">The Body for the Email</param>
        /// <param name="isBodyHtml">Value to indicate whether the Body is HTML or not, defaults to True</param>
        /// <param name="filePath">The Path to the Generated Report, if left Empty, it will generate the report itself</param>
        /// <returns>True if successfully sent to ALL recipients, False if not</returns>
        /// <exception cref="ArgumentException">
        /// Throws an <see cref="ArgumentNullException"/> if there are 0 
        /// Email Recipients Setup on the <see cref="GenericReport"/>
        /// </exception>
        public static bool SendReportResultsToEmails(ReportResultSet result, 
            string subject, 
            string body,
            bool isBodyHtml = true,
            string filePath = null)
        {
            var report = result.ReportUsed;
            if (!report.EmailRecipients.Any())
            {
                throw new ArgumentException("Email Recipients was empty.");
            }

            var successCount = 0;

            try
            {
                filePath = SmtpEmailSender.HandleFilePath(filePath, result);
                var fileName = filePath.Substring(filePath.LastIndexOf('\\') + 1);

                var attachment = new Attachment(filePath)
                {
                    Name = fileName
                };

                report.EmailRecipients.ForEach(recipient =>
                {
                    var mailMessage = new MailMessage(SmtpEmailSender._fromAddress, 
                        recipient, 
                        subject,
                        body)
                    {
                        IsBodyHtml = isBodyHtml
                    };

                    //// Add attachment for the report
                    mailMessage.Attachments.Add(attachment);

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