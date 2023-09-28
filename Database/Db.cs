using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using RpaLib.ProcessAutomation;
using OpenQA.Selenium;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;

namespace RpaLib.Database
{
    public enum SqlCommandType
    {
        NotDetected,
        Select,
        CreateTable,
        CreateDatabase,
        Insert,
        Update,
        Delete,
        Drop,
    }

    public static class Db
    {
        public static QueryReturn DataReaderToDataTable(DbDataReader reader, bool debugMessages = false)
        {
            QueryReturn queryReturn = new QueryReturn();
            // if the query returned something
            // add the column info to the datatable obj
            if (reader.HasRows)
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    DataColumn col = new DataColumn();
                    col.DataType = reader.GetFieldType(i);
                    col.ColumnName = reader.GetName(i);

                    queryReturn.Table.Columns.Add(col);

                    if (debugMessages)
                    {
                        Trace.WriteLine(string.Join(Environment.NewLine,
                        $"Column: \"{reader.GetName(i)}\"",
                        $"    -> DB Datatype:     \"{reader.GetDataTypeName(i)}\"",
                        $"    -> DotNet Datatype: \"{reader.GetFieldType(i)}\""));
                    }
                }
            }

            for (int i = 0; reader.Read(); i++)
            {
                // Creates a new row with the result columns' info
                DataRow row = queryReturn.Table.NewRow();
                queryReturn.Table.Rows.Add(row);

                for (int j = 0; j < reader.FieldCount; j++)
                {
                    queryReturn.Table.Rows[i][j] = reader.GetValue(j);
                }
            }

            return queryReturn;
        }

        public static SqlCommandType DetectSqlCommandType(string sqlCommand)
        {

            SqlCommandType sqlCommandType = SqlCommandType.NotDetected;

            if (Rpa.IsMatch(sqlCommand, @"^\s*SELECT", RegexOptions.IgnoreCase))
            {
                sqlCommandType = SqlCommandType.Select;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*CREATE\s*TABLE"))
            {
                sqlCommandType = SqlCommandType.CreateTable;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*CREATE\s*DATABASE"))
            {
                sqlCommandType = SqlCommandType.CreateDatabase;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*INSERT\s*INTO"))
            {
                sqlCommandType = SqlCommandType.Insert;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*UPDATE"))
            {
                sqlCommandType = SqlCommandType.Update;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*DELETE"))
            {
                sqlCommandType = SqlCommandType.Delete;
            }
            else if (Rpa.IsMatch(sqlCommand, @"^\s*DROP"))
            {
                sqlCommandType = SqlCommandType.Drop;
            }

            return sqlCommandType;
        }

        public static dynamic[] DatatableToArray(DataTable datatable)
        {
            List<dynamic> lines = new List<dynamic>();

            foreach (DataRow row in datatable.Rows)
            {
                List<dynamic> columns = new List<dynamic>();
                foreach (DataColumn col in datatable.Columns)
                {
                    columns.Add(row[col.ColumnName]);
                }
                lines.Add(columns.ToArray());
            }

            return lines.ToArray();
        }
    }
}
