using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO;
using sapfewse;
using System.Diagnostics;
using RpaLib.Tracing;
using System.Globalization;
using RpaLib.ProcessAutomation;

/// General Robot oriented class to ease manipulation of SQLite
/// and possibly other DBs in future
namespace RpaLib.Database
{
    public class Sqlite
    {
        private string _connectionString;
        private string _dbFilePath;
        private bool _debugMessages = false;
 
        public Sqlite(string dbFilePath)
        {
            _dbFilePath = Path.GetFullPath(dbFilePath);
            ConnectionString = dbFilePath;
        }
        public Sqlite(string dbFilePath, string creationTablesDirPath) : this(dbFilePath)
        {
            CreatePopulateTablesIfDbNotExists(creationTablesDirPath);
        }

        public string ConnectionString
        {
            get
            {
                return this._connectionString;
            }
            set
            {
                _connectionString = string.Concat("Data Source=", _dbFilePath);
            }
        }
        public QueryReturn QueryResult { get; private set; }

        /// <summary>
        /// Subclass with the return type of any Query
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
                return String.Format(@"
            AffectedRows: {0}; TableRows: {1}
            ", AffectedRows, Table.Rows.Count);
            }
        }

        public string Quote(dynamic objToQuote)
        {
            if (objToQuote == null)
                return "''";
            else if (objToQuote.GetType() == typeof(string))
                return string.Format("'{0}'", Regex.Replace(objToQuote, "'", @"''"));
            else if (objToQuote.GetType() == typeof(DateTime))
                return Quote(objToQuote.ToString("yyyy-MM-dd"));
            else
                return string.Format(CultureInfo.InvariantCulture, "{0}", objToQuote);
        }

        public string RmQuotes(string s)
        {
            Debug($"Removing quotes from value \"{s}\"");
            if (s == null)
                return string.Empty;
            else
                return s.Replace("'", string.Empty);
        }

        public QueryReturn Query(string sqlcmd)
        {
            QueryResult = new QueryReturn();
            Log.Write($"Trying to run the database query:\n{sqlcmd}");
            using (var connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = sqlcmd;

                if (Regex.IsMatch(sqlcmd, @"^\s*SELECT", RegexOptions.IgnoreCase))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        // if the query returned something
                        // add the column info to the datatable obj
                        if (reader.HasRows)
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                DataColumn col = new DataColumn();
                                col.DataType = reader.GetFieldType(i);
                                col.ColumnName = reader.GetName(i);

                                QueryResult.Table.Columns.Add(col);

                                if (_debugMessages)
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
                            DataRow row = QueryResult.Table.NewRow();
                            QueryResult.Table.Rows.Add(row);

                            for (int j = 0; j < reader.FieldCount; j++)
                            {
                                QueryResult.Table.Rows[i][j] = reader.GetValue(j);
                            }
                        }
                    }
                    Console.WriteLine($"The query result is: \"{QueryResult}\"");
                }
                else
                {
                    if (Regex.IsMatch(sqlcmd, @"^\s*CREATE\s*TABLE", RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine("Command: CREATE TABLE");
                    }
                    else if (Regex.IsMatch(sqlcmd, @"^\s*INSERT\s*INTO", RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine("Command: INSERT INTO");
                    }
                    QueryResult.AffectedRows = command.ExecuteNonQuery();
                    Console.WriteLine($"Executed non-query. Rows affected: {QueryResult.AffectedRows}");
                }
            }
            Log.Write("DB query: success.");
            return QueryResult;
        }

        public DataTable Select(string query) => Query(query).Table;

        public DataRow SelectFirstRow(string query)
        {
            DataTable queryResult = Select(query);
            if (queryResult.Rows.Count == 0) return null;
            else return queryResult.Rows[0];
        }

        public void CreatePopulateTablesIfDbNotExists(string sqlFilesDirPath)
        {
            //Rpa.MessageBox("Checking");
            if (!File.Exists(Path.GetFullPath(this._dbFilePath)))
            {
                //Rpa.MessageBox("will create");
                this.CreatePopulateTables(sqlFilesDirPath);
            }
        }
        public void CreatePopulateTables(string sqlFilesDirPath)
        {
            foreach (string file in Directory.GetFiles(Path.GetFullPath(sqlFilesDirPath)))
            {
                if (Regex.IsMatch(file, @".+\.sql", RegexOptions.IgnoreCase))
                {
                    string sqlCreationCmd = File.ReadAllText(file);
                    this.Query(sqlCreationCmd);
                }
            }
        }

        public void ResetDb(string sqlFilesDirPath)
        {
            QueryReturn sqliteSchemaTables = this.Query(string.Join(" ",
                "SELECT name",
                "FROM sqlite_schema",
                "WHERE type='table'",
                "AND name NOT LIKE 'sqlite_%'"
            ));

            foreach (DataRow row in sqliteSchemaTables.Table.Rows)
            {
                this.Query($"DROP TABLE {row["name"]}");
            }

            this.CreatePopulateTables(sqlFilesDirPath);
        }

        private void Debug(string message)
        {
            if (_debugMessages)
                Trace.WriteLine(message);
        }
    }
}
