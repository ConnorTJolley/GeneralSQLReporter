namespace GeneralSQLReporter.Models
{
    using System;

    /// <summary>
    /// Interaction Logic for the <see cref="SqlRowValue"/> Class.
    /// </summary>
    /// <remarks>
    /// Contained within the <see cref="SqlRow.Values"/> Property
    /// </remarks>
    public class SqlRowValue
    {
        /// <summary>
        /// Gets the Index for the Row in relation to the results from the <see cref="GenericReport"/>.
        /// </summary>
        public int RowNumber { get; internal set; }

        /// <summary>
        /// Gets the Index for the Column in relation to the results from the <see cref="GenericReport"/>.
        /// </summary>
        public int ColumnIndex { get; internal set; }

        /// <summary>
        /// Gets the Value that was retrieved from the results fro the <see cref="GenericReport"/>
        /// </summary>
        /// <remarks>
        /// Used in Conjunction with the <see cref="SqlRowValue.DataType"/> Property to convert to it's representative form.
        /// </remarks>
        public object Value { get; internal set; }

        /// <summary>
        /// Gets the <see cref="Type"/> for the <see cref="SqlRowValue.Value"/> object
        /// </summary>
        public Type DataType { get; internal set; }
    }
}