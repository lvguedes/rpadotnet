using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.IO;
using System.Globalization;
using RpaLib.Tracing;


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
                return _connectionString;
            }
            set
            {
                _connectionString = string.Concat("Data Source=", _dbFilePath);
            }
        }
        public QueryReturn QueryResult { get; private set; }

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

                var sqlCommandType = Db.DetectSqlCommandType(sqlcmd);

                if (sqlCommandType == SqlCommandType.Select)
                {
                    QueryResult = Db.DataReaderToDataTable(command.ExecuteReader());
                    Trace.WriteLine($"The query result is: \"{QueryResult}\"");
                }
                else
                {
                    Trace.WriteLine($"Command: \"{sqlCommandType}\"");
                    QueryResult.AffectedRows = command.ExecuteNonQuery();
                    Trace.WriteLine($"Executed non-query. Rows affected: {QueryResult.AffectedRows}");
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
            if (!File.Exists(Path.GetFullPath(_dbFilePath)))
            {
                //Rpa.MessageBox("will create");
                CreatePopulateTables(sqlFilesDirPath);
            }
        }
        public void CreatePopulateTables(string sqlFilesDirPath)
        {
            foreach (string file in Directory.GetFiles(Path.GetFullPath(sqlFilesDirPath)))
            {
                if (Regex.IsMatch(file, @".+\.sql", RegexOptions.IgnoreCase))
                {
                    string sqlCreationCmd = File.ReadAllText(file);
                    Query(sqlCreationCmd);
                }
            }
        }

        public void ResetDb(string sqlFilesDirPath)
        {
            QueryReturn sqliteSchemaTables = Query(string.Join(" ",
                "SELECT name",
                "FROM sqlite_schema",
                "WHERE type='table'",
                "AND name NOT LIKE 'sqlite_%'"
            ));

            foreach (DataRow row in sqliteSchemaTables.Table.Rows)
            {
                Query($"DROP TABLE {row["name"]}");
            }

            CreatePopulateTables(sqlFilesDirPath);
        }

        private void Debug(string message)
        {
            if (_debugMessages)
                Trace.WriteLine(message);
        }
    }
}
