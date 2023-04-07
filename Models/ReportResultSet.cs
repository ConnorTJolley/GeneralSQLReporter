namespace GeneralSQLReporter.Models
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Data.Common;

    public class ReportResultSet
    {
        public SqlReport ReportUsed { get; internal set; }
        public ReadOnlyCollection<DbColumn> Columns { get; internal set; }
        public List<SqlRow> Rows { get; set; }
    }
}