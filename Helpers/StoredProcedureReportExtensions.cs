namespace GeneralSQLReporter.Helpers
{
    using System.Data.SqlClient;
    using System.Data;
    using GeneralSQLReporter.Models;
    using System.Linq;

    /// <summary>
    /// Extensions Methods for the <see cref="StoredProcedureReport"/> Classs
    /// </summary>
    /// <remarks>
    /// Main functionality is to provide easy methods to add elements such as Parameters to the Stored Procedure
    /// as well as creating a test output SQL command for test purposes.
    /// </remarks>
    public static class StoredProcedureReportExtensions
    {
        /// <summary>
        /// Adds a Named Parameter with a <see cref="SqlDbType"/> and Value to the <see cref="StoredProcedureReport"/>
        /// </summary>
        /// <param name="report">The <see cref="StoredProcedureReport"/> to Extend and Add the Parameter to</param>
        /// <param name="parameterName">The Name of the Parameter, e.g @accountNumber</param>
        /// <param name="type">The <see cref="SqlDbType"/> For the Parameter, e.g <see cref="SqlDbType.NVarChar"/></param>
        /// <param name="value">The Value to assign to the Parameter</param>
        public static void AddParameter(this StoredProcedureReport report,
            string parameterName,
            SqlDbType type,
            object value) =>
                report.Parameters.Add(new SqlParameter(parameterName, type)
                {
                    Value = value
                });

        /// <summary>
        /// Handles Checking of the <see cref="StoredProcedureReport"/> Is a Valid report or not
        /// </summary>
        /// <param name="report">The <see cref="StoredProcedureReport"/> to extend and Check</param>
        /// <returns>True if valid, False if not</returns>
        public static bool IsValid(this StoredProcedureReport report)
        {
            if (string.IsNullOrWhiteSpace(report.ProcedureName.Trim()) ||
                (report.EmailRecipients.Any() && !SmtpEmailSender.IsSetup()))
            {
                return false;
            }

            return true;
        }
    }
}