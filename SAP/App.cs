using Microsoft.Office.Interop.Excel;
using RpaLib.ProcessAutomation;
using RpaLib.SAP.Exceptions;
using RpaLib.Tracing;
using sapfewse;
using saprotwr.net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    /// <summary>
    /// Class to verify info, get objects directly from native COM object
    /// </summary>
    public class App
    {
        private const int _getSapObjTimeoutSeconds = 20;
        public GuiApplication GuiApplication { get; private set; }
        public string ConnectionErrorText { get; private set; }
        public Connection[] Connections { get; private set; }

        public App ()
        {
            Update();
        }

        /// <summary>
        /// Low-level function to get the SAP Interop object
        /// </summary>
        /// <param name="timeout">Seconds to stay retrying to get the SAP object when the process is still starting.</param>
        /// <returns></returns>
        public static GuiApplication GetSapInteropApp(int timeout = _getSapObjTimeoutSeconds)
        {
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object sapGuilRot = null;
            int timePassed = 0;

            // try to get SapGuilRot within timeout
            while (sapGuilRot == null && timePassed < timeout)
            {
                sapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
                Thread.Sleep(1000);
                timePassed++;
            }

            // use GuilRot to get the scripting engine
            object engine = sapGuilRot.GetType().InvokeMember(
                "GetSCriptingEngine",
                System.Reflection.BindingFlags.InvokeMethod,
                null, sapGuilRot, null);

            return engine as GuiApplication;
        }

        /// <summary>
        /// After creating a new session or connection you must call this method so that this
        /// class' object re-fetch the new content created from the COM object. Otherwise, the
        /// new session or connection items won't appear in this class properties.
        /// </summary>
        /// <returns>This class' updated object</returns>
        public App Update()
        {
            GuiApplication = GetSapInteropApp();
            ConnectionErrorText = GuiApplication.ConnectionErrorText;
            Connections = GetConnections();

            return this;
        }

        /// <summary>
        /// Convert a GuiCollection of GuiConnection to a string of Connection.
        /// Connection is a wrapper to GuiConnection.
        /// </summary>
        /// <returns>An array of Connection objects.</returns>
        public static Connection[] GetConnections()
        {
            var connections = new List<Connection>();
            foreach (GuiConnection conn in GetSapInteropApp().Connections)
            {
                connections.Add(new Connection(conn));
            }

            return connections.ToArray();
        }

        /// <summary>
        /// Find a connection using Regex pattern over its path ID.
        /// </summary>
        /// <param name="connectionIdPattern">Regex pattern to be searched in Connection's Id property.</param>
        /// <returns>The Connection object found or null.</returns>
        public static Connection FindConnectionById(string connectionIdPattern)
        {
            var selectedConn = from conn in GetConnections()
                               where Rpa.IsMatch(conn.Id, connectionIdPattern)
                               select conn;

            return selectedConn.FirstOrDefault();
        }

        public static Session[] FindSessions(string transactionNameRegex)
        {
            var openedConnections = App.GetConnections();

            if (openedConnections.Length != 1)
                throw new ConnectionNotUniqueException(openedConnections.Length);

            var currentConnection = openedConnections[0];

            var sessions = from s in currentConnection.Sessions
                           where Rpa.IsMatch(s.Transaction, transactionNameRegex)
                           select s;

            return sessions.ToArray();
        }

        /// <summary>
        /// Returns a string containing info about all connections
        /// </summary>
        /// <returns>string with info about all connections</returns>
        public static string ConnectionsInfo()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var connection in GetConnections())
            {
                if (!string.IsNullOrEmpty(sb.ToString()))
                    sb.Append("\n    ,\n");
                sb.Append(connection);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Prints on screen (Trace stream) info about all connections fetched from COM object. 
        /// </summary>
        public static void ShowConnectionsInfo()
        {
            Trace.WriteLine(ConnectionsInfo());
        }

        public static void CloseAll()
        {
            foreach (var connection in GetConnections())
            {
                foreach (var session in connection.Sessions)
                    session.Close();
            }
        }
    }
}
