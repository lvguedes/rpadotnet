using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using RpaLib.Tracing;
using RpaLib.ProcessAutomation;
using System.Web.UI.WebControls;

namespace RpaLib.Database
{
    public class SharePoint : IDisposable
    {
        public OleDbConnection Connection { get; private set; }
        public string ConnString { get; private set; }
        public QueryReturn QueryReturn { get; private set; }

        public SharePoint(string connectionString)
        {
            ConnString = connectionString;
            Connection = new OleDbConnection(connectionString);
            TryConnect(Connection);
        }

        // sqlParameters.Key must be the parameter name. Example: @age
        // sqlParameters.Value must be the parameter value. Example: 18
        public QueryReturn Query(string sqlCommand, Dictionary<string, dynamic> sqlParameters = null, bool exclusiveConnection = false)
        {
            OleDbConnection conn = Connection;

            if (exclusiveConnection)
            {
                conn = new OleDbConnection(ConnString);
                TryConnect(conn);
            }

            var sqlCommandType = Db.DetectSqlCommandType(sqlCommand);
            OleDbCommand command = new OleDbCommand(sqlCommand, conn);

            foreach (var param in sqlParameters?? new Dictionary<string, dynamic>())
            {
                command.Parameters.AddWithValue(param.Key, param.Value);
            }

            if (sqlCommandType == SqlCommandType.Select)
            {
                OleDbDataReader reader = command.ExecuteReader();
                QueryReturn = Db.DataReaderToDataTable(reader);
            }
            else
            {
                Trace.WriteLine($"Running SQL command: \"{sqlCommandType}\"");
                QueryReturn.AffectedRows = command.ExecuteNonQuery();
                Trace.WriteLine($"Executed non-query. Rows affected: {QueryReturn.AffectedRows}");
            }

            if (exclusiveConnection)
            {
                conn.Close();
            }

            return QueryReturn;
        }

        private void TryConnect(OleDbConnection connection)
        {
            try
            {
                connection.Open();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Error when trying to connect to OLE.DB Sharepoint.", color: ConsoleColor.Red);
                Trace.WriteLine(ex);
                Console.ReadKey();
                Environment.Exit(1);
            }
        }

        public void Dispose()
        {
            Connection.Close();
        }

        ~SharePoint()
        {
            Dispose();
        }
    }
}
