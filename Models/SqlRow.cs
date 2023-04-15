namespace GeneralSQLReporter.Models
{
    using System.Collections.Generic;

    /// <summary>
    /// Interaction Logic for the <see cref="SqlRow"/> Class.
    /// </summary>
    /// <remarks>
    /// Used in the <see cref="ReportResultSet.Rows"/> Property
    /// </remarks>
    public class SqlRow
    {
        /// <summary>
        /// Gets the Index for the Row in relation to the results from the <see cref="GenericReport"/>.
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// Gets the <see cref="List{SqlRowValue}"/> for the <see cref="SqlRow"/>
        /// </summary>
        public List<SqlRowValue> Values { get; internal set; } = new List<SqlRowValue>();

        /// <summary>
        /// Initializes a new Instance of the <see cref="SqlRow"/> Class.
        /// </summary>
        /// <param name="index">The Index for the Row in relation to the results from the <see cref="GenericReport"/></param>
        internal SqlRow(int index)
        {
            this.Index = index;
        }
    }
}