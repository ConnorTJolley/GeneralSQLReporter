namespace GeneralSQLReporter.Models
{
    /// <summary>
    /// Enumerator for the Different Report Types
    /// </summary>
    public enum ReportType
    {
        /// <summary>
        ///  Unknown not determined
        /// </summary>
        Unknown = -1,

        /// <summary>
        /// Indicates that it is a <see cref="SqlReport"/>
        /// </summary>
        SqlReport = 0,

        /// <summary>
        /// Indicates that it is a <see cref="StoredProcedureReport"/>
        /// </summary>
        StoredProcedure = 1
    }
}