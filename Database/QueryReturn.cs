using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using RpaLib.ProcessAutomation;

namespace RpaLib.Database
{
    /// <summary>
    /// Class that represents the return type of any Query
    /// </summary>
    public class QueryReturn
    {
        public QueryReturn()
        {
            Table = new DataTable();
        }
        public int AffectedRows { get; set; }
        public DataTable Table { get; private set; }

        public override string ToString()
        {
            return string.Format(@"
            AffectedRows: {0}; TableRows: {1}
            ", AffectedRows, Table.Rows.Count);
        }

        public string ToJson()
        {
            return Rpa.Json(ToArray());
        }

        public dynamic[] ToArray()
        {
            return Db.DatatableToArray(Table?? new DataTable());
        }
    }
}
