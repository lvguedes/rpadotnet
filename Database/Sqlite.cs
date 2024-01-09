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
using RpaLib.ProcessAutomation;


/// General Robot oriented class to ease manipulation of SQLite
/// and possibly other DBs in future
namespace RpaLib.Database
{
    public class Sqlite
    {
        private string _connectionString;
        private string _dbFilePath;
        public bool DebugMessages { get; }
 
        public Sqlite(string dbFilePath, bool debugMessages = false)
        {
            _dbFilePath = Ut.GetFullPath(dbFilePath);
            ConnectionString = dbFilePath;
            DebugMessages = debugMessages;
        }
        public Sqlite(string dbFilePath, string creationTablesDirPath, bool debugMessages = false)
            : this(dbFilePath, debugMessages)
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

        public QueryReturn Query(string sqlcmd, bool debugMessages = false)
        {
            QueryResult = new QueryReturn();
            Debug($"Trying to run the database query:\n{sqlcmd}");
            using (var connection = new SQLiteConnection(ConnectionString))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = sqlcmd;

                var sqlCommandType = Db.DetectSqlCommandType(sqlcmd);

                if (sqlCommandType == SqlCommandType.Select)
                {
                    QueryResult = Db.DataReaderToDataTable(command.ExecuteReader(), debugMessages);
                    Debug($"The query result is: \"{QueryResult}\"");
                }
                else
                {
                    Debug($"Command: \"{sqlCommandType}\"");
                    QueryResult.AffectedRows = command.ExecuteNonQuery();
                    Debug($"Executed non-query. Rows affected: {QueryResult.AffectedRows}");
                }
            }
            Debug("DB query: success.");
            return QueryResult;
        }

        public QueryReturn[] QueryAll(params string[] queries)
        {
            return QueryAll(false, queries);
        }

        public QueryReturn[] QueryAll(bool debugMessages, params string[] queries)
        {
            List<QueryReturn> queryResult = new List<QueryReturn>();
            foreach (var query in queries)
            {
                queryResult.Add(Query(query, debugMessages));
            }

            return queryResult.ToArray();
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

        private void Debug(string message, bool debugMessages = false)
        {
            if (DebugMessages || debugMessages)
                Trace.WriteLine(message);
        }
    }
}
