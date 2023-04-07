namespace GeneralSQLReporter.Models
{
    public class SqlRow
    {
        public int RowNumber { get; set; }
        public int ColumnIndex { get; set; }
        public object Value { get; set; }
    }
}