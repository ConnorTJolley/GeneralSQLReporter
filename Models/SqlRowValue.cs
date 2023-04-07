using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneralSQLReporter.Models
{
    public class SqlRowValue
    {
        public int RowNumber { get; internal set; }
        public int ColumnIndex { get; internal set; }
        public object Value { get; internal set; }
        public Type DataType { get; internal set; }
    }
}
