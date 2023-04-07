namespace GeneralSQLReporter.Models
{
    using System.Data;

    public class SqlColumn
    {
        public int Index { get; internal set; }
        public string Name { get; internal set; }
        public SqlDbType DataType { get; internal set; }
    }
}