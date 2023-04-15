namespace GeneralSQLReporter.Models
{
    /// <summary>
    /// Interaction Logic for the <see cref="SqlColumn"/> Class.
    /// </summary>
    /// <remarks>
    /// Used in the <see cref="ReportResultSet.Columns"/> Property
    /// </remarks>
    public class SqlColumn
    {
        /// <summary>
        /// Gets the Column Index returned by the Schema
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// Gets the Column Name returned by the Schema
        /// </summary>
        public string Name { get; internal set; }
    }
}