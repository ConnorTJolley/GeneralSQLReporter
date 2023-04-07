namespace GeneralSQLReporter.Models
{
    using System.Collections.Generic;
    using System.Data;

    public class ReportResultSet
    {
        public SqlReport ReportUsed { get; internal set; }
        public List<SqlColumn> Columns { get; internal set; } = new List<SqlColumn>();
        public List<SqlRow> Rows { get; set; } = new List<SqlRow>();
    }
}