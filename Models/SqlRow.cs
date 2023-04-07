using System;
using System.Collections.Generic;

namespace GeneralSQLReporter.Models
{
    public class SqlRow
    {
        public int Index { get; internal set; }
        public List<SqlRowValue> Values { get; internal set; } = new List<SqlRowValue>();

        public SqlRow(int index)
        {
            this.Index = index;
        }
    }
}