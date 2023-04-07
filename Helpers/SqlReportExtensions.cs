namespace GeneralSQLReporter.Helpers
{
    using System.Linq;
    using GeneralSQLReporter.Models;

    public static class SqlReportExtensions
    {
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