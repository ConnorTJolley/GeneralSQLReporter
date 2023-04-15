namespace GeneralSQLReporter.Helpers
{
    using System.Linq;
    using GeneralSQLReporter.Models;

    /// <summary>
    /// Extension Methods for the <see cref="SqlReport"/> Class
    /// </summary>
    public static class SqlReportExtensions
    {
        /// <summary>
        /// Determines if a <see cref="SqlReport"/> is Valid or Not
        /// </summary>
        /// <param name="report">The <see cref="SqlReport"/> to extend and check.</param>
        /// <returns>True if Valid, False if not 
        /// (or if <see cref="GenericReport.EmailRecipients"/> has records but <see cref="SmtpEmailSender"/> is not setup)</returns>
        public static bool IsValid(this SqlReport report)
        {
            if (string.IsNullOrWhiteSpace(report.Query.Trim()) || 
                (report.EmailRecipients.Any() && !SmtpEmailSender.IsSetup()))
            {
                return false;
            }

            return true;
        }
    }
}