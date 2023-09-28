using RpaLib.Tracing;
using RpaLib.ProcessAutomation;
using sapfewse;
using saprotwr.net;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Data;

namespace RpaLib.SAP
{
    public static class Sap
    {
        private static readonly int connTimeoutSeconds = 10;
        private static GuiConnection _connection;
        private static Session _session;
        public static Session Session
        {
            get => _session;
            set
            {
                _session = value;
                Tracing.Log.Write(string.Join(Environment.NewLine,
                    "Current session (Session) set to:",
                    Session.SessionInfo(value.GuiSession)));
            }
        }
        public static Session[] Sessions { get; set; }
        public static int TriedConnect { get; private set; } = 0;
        public static GuiApplication App { get; set; }
        public static GuiConnection Connection
        {
            get => _connection;
            set
            {
                _connection = value;
                Tracing.Log.Write(string.Join(Environment.NewLine,
                    $"Current connection (Connection) set to:",
                    ConnectionInfo(value)));
            }
        }
        public static GuiConnection[] Connections { get; set; }
        public static dynamic SapExe { get; } = new
        {
            BaseName = @"saplogon",
            FullPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        };

        public enum StatusType
        {
            Error,
            Warning,
            Success,
            Abort,
            Information
        }

        private static string GetStatusTypeLetter(StatusType statusType)
        {
            switch(statusType)
            {
                case StatusType.Error: return "E";
                case StatusType.Warning: return "W";
                case StatusType.Success: return "S";
                case StatusType.Abort: return "A";
                case StatusType.Information: return "I";
                default: return null;
            }
        }

        private static GuiApplication GetSapInteropApp()
        {
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            object engine = SapGuilRot.GetType().InvokeMember(
                "GetSCriptingEngine",
                System.Reflection.BindingFlags.InvokeMethod,
                null, SapGuilRot, null);

            return engine as GuiApplication;
        }

        private static void CreateSapConnection(string connectionName)
        {
            App = GetSapInteropApp();

            if (App.Connections.Count == 0)
            {
                Connection = App.OpenConnection(connectionName);
            }
            else
            {
                Connection = App.Connections.ElementAt(0) as GuiConnection;
            }
        }

        public static void Connect(string connectionName, int tryingsLimitInSeconds)
        {
            try
            {
                Tracing.Log.Write($"Connecting with SAP: \"{connectionName}\"");
                CreateSapConnection(connectionName);
                MapExistingSessions();
                Tracing.Log.Write($"Connection Succeeded.\n{ConnectionInfo()}");
            }
            // exception thrown when SAP Connection window is not opened.
            catch (NullReferenceException ex)
            {
                Tracing.Log.Write($"Caught exception\n{ex}");
                // retry connection within config's file limit
                TriedConnect++;
                if (TriedConnect <= tryingsLimitInSeconds)
                {
                    if (TriedConnect == 1)
                    {
                        Tracing.Log.Write($"Killing SAP processes by name if it exists: {SapExe.BaseName}");
                        Rpa.KillProcess(SapExe.BaseName);

                        Tracing.Log.Write($"Starting SAP exe: {SapExe.FullPath}");
                        Rpa.StartWaitProcess(SapExe.FullPath, outputProcesses: true);
                    }

                    Tracing.Log.Write($"Tried connecting to SAP for: {TriedConnect} times. Trying again.");
                    Thread.Sleep(1000);
                    Connect(connectionName, tryingsLimitInSeconds);
                }
                else
                {
                    throw new ExceededRetryLimitSapConnectionException();
                }
            }
        }

        // create a new connection or uses the current actually? (if it creates this is an error)
        public static void UpdateConnections()
        {
            List<GuiConnection> connections = new List<GuiConnection>();
            App = GetSapInteropApp();

            foreach (GuiConnection conn in App.Connections)
            {
                connections.Add(conn);
            }
            Connections = connections.ToArray();

            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Updating connections...",
                $"The number of existing connections: {connections.Count}",
                $"Available connections:",
                ConnectionsInfo()));

            Connection = App.Connections.ElementAt(0) as GuiConnection;
            
        }

        public static void MapExistingSessions()
        {
            List<Session> sessions = new List<Session>();

            Tracing.Log.Write($"Trying to map existing sessions. Number of existing sessions: {Connection.Sessions.Count}");

            foreach (GuiSession session in Connection.Sessions)
            {
                Tracing.Log.Write(string.Join(Environment.NewLine,
                    "Mapping session:",
                    Session.SessionInfo(session)));
                sessions.Add(new Session(session));
            }

            // sort by Session.GuiSession.Info.SessionNumber
            Sessions = sessions.OrderBy(x => x.GuiSession.Info.SessionNumber).ToArray();
        }

        // high-level function
        public static Session CreateNewSession(string connectionName, string transactionId = null, int useSessionId = -1, int connectionTimeoutSeconds = 10)
        {
            Session session;

            if (Sap.Connection == null)
            {
                Sap.Connect(connectionName, connectionTimeoutSeconds);
            }
            if (useSessionId >= 0)
            {
                // to work with specific session id
                session = Sap.Sessions[0];
            }
            else
            {
                // to create a new session and work upon it
                session = Sap.CreateNewSession();
            }
            session.GuiSession.LockSessionUI();
            session.FindById<GuiFrameWindow>("wnd[0]").Iconify();
            if (!string.IsNullOrEmpty(transactionId))
                session.AccessTransaction(transactionId);

            return session;
        }

        // low-level function
        private static Session CreateNewSession()
        {
            int lastSession = Sap.Sessions.Length - 1;
            Sap.Sessions[lastSession].CreateNewSession();
            return FindSession("SESSION_MANAGER");
        }

        public static Session FindSession(string transaction) => Sessions.Where(x => x.GuiSession.Info.Transaction.Equals(transaction)).FirstOrDefault();

        public static void AccessTransaction(string id) => AccessTransaction(Session, id);
        public static void AccessTransaction(Session session, string id)
        {
            Tracing.Log.Write($"Trying to access transaction \"{id}\"");
            session.GuiSession.StartTransaction(id);
            Tracing.Log.Write($"Transaction Access: Successful. Transaction assigned to session: [{session.Index}]");
        }

        public static string ConnectionInfo() => ConnectionInfo(Connection);
        public static string ConnectionInfo(GuiConnection connection)
        {
            return string.Join(Environment.NewLine,
                $"  Connection:",
                $"    Description: \"{connection.Description}\"",
                $"    ConnectionString: \"{connection.ConnectionString}\"",
                $"    Sessions: \"{connection.Sessions}\"",
                $"    Children: \"{connection.Children}\"",
                $"    DisabledByServer: \"{connection.DisabledByServer}\"",
                $"",
                $"  The Sessions/Children elements:",
                SessionsInfo()
                );
        }

        public static string ConnectionsInfo()
        {
            StringBuilder info = new StringBuilder();

            for (int i = 0; i < Connections.Length; i++)
            {
                info.AppendLine(string.Join(Environment.NewLine,
                    $"[{i}]",
                    ConnectionInfo(Connections[i])
                    ));
            }
            return info.ToString();
        }

        public static string SessionsInfo()
        {
            StringBuilder info = new StringBuilder();
            if (Sessions == null) return "No Sessions";
            foreach (Session session in Sessions)
            {
                info.AppendLine(string.Concat(Environment.NewLine, session));
            }

            return info.ToString();
        }

        public static string MainWindowInfo()
        {
            GuiMainWindow mw = Session.FindById("wnd[0]") as GuiMainWindow;
            return string.Join(Environment.NewLine,
                $"GuiMainWindow info:",
                $"  ButtonbarVisible: {mw.ButtonbarVisible}",
                $"  StatusbarVisible: {mw.StatusbarVisible}",
                $"  TitlebarVisible:  {mw.TitlebarVisible}",
                $"  ToolbarVisible:   {mw.ToolbarVisible}"
                );
        }

        public static string AllSessionIdsInfo() => AllSessionIdsInfo(Session.GuiSession);

        public static string AllSessionIdsInfo(GuiSession session)
        {
            return AllDescendantIdsInfo(session);
        }

        public static string AllDescendantIdsInfo(dynamic guiContainer)
        {
            Dictionary<string, dynamic>[] ids = AllDescendantIds(guiContainer);
            StringBuilder info = new StringBuilder($"All IDs from \"[{guiContainer.Type}]: {guiContainer.Id}\" to its innermost descendant are:\n");

            foreach (var id in ids)
            {
                info.AppendLine($"[{id["Type"]}]: {id["PathId"]}");
            }

            return info.ToString();
        }

        public static T[] FindByText<T>(GuiSession session, string labelText)
        {
            T[] objFound = AllSessionIds(session)
                .Cast<Dictionary<string, dynamic>>()
                .Where(d =>
                {
                    GuiVComponent visualComponent;
                    try
                    {
                        visualComponent = (GuiVComponent)d["Obj"];
                        return Rpa.IsMatch(visualComponent.Text, labelText);
                    }
                    catch (InvalidCastException)
                    {
                        // if not a visual component, there isn't text
                        return false;
                    }
                })
                //.Select( d => (T)d["Obj"] )
                ///*
                .Select( (d) => {
                    Log.Write($"Converting to {typeof(T)} id: {d["PathId"]}");
                    try
                    {
                        return (T)d["Obj"];
                    }
                    catch (InvalidCastException)
                    {
                        return default(T);
                    }
                })
                .Where(x => x != null )
                //*/
                .ToArray();

            /* 
            // Not good when quering the result
            if (objFound.Count() == 0)
                objFound = null;
            */

            return objFound;
        }

        public static Dictionary<string, dynamic>[] AllSessionIds(GuiSession session) => AllDescendantIds(session);
        public static Dictionary<string, dynamic>[] AllDescendantIds(dynamic root)
        {
            List<Dictionary<string,dynamic>> ids = new List<Dictionary<string,dynamic>>();

            AllDescendantIds(ids, root);

            return ids.ToArray();
        }
        private static List<Dictionary<string, dynamic>> AllDescendantIds(List<Dictionary<string, dynamic>> ids, dynamic root)
        {
            /*
            foreach (dynamic child in root.Children)
            {
                if (child.ContainerType)
                    AllDescendantIds(ids, child);
            }
            */

            Action<dynamic> addNodeToList =
                (dynamic node) =>
                {
                    ids.Add(
                        new Dictionary<string, dynamic>
                        {
                            { "PathId", (string)node.Id },
                            { "Type", node.Type }, // Type prop from GuiComponent
                            { "Obj", node}
                        });
                };

            for (int i = 0; i < root.Children.Count; i++)
            {
                dynamic child = root.Children.ElementAt[i];
                if (child.ContainerType)
                    AllDescendantIds(ids, child);
                else
                    addNodeToList(child);
            }

            addNodeToList(root);

            return ids;
        }

        public static string CurrentStatusBarInfo()
        {
            GuiStatusbar currentStatusBar = Sap.Session.FindById<GuiStatusbar>("wnd[0]/sbar");

            return string.Join(Environment.NewLine,
                $"Current Status Bar properties:",
                $"  MessageType: {currentStatusBar.MessageType}",
                $"  Text: {currentStatusBar.Text}");
        }
        public static bool IsStatusType(StatusType status)
        {
            string statusLetter = GetStatusTypeLetter(status);
            if (Regex.IsMatch(Sap.Session.FindById<GuiStatusbar>("wnd[0]/sbar").MessageType, statusLetter, RegexOptions.IgnoreCase))
                return true;
            else
                return false;
        }

        public static bool IsStatusMessage(string message)
        {
            if (Regex.IsMatch(Sap.Session.FindById<GuiStatusbar>("wnd[0]/sbar").Text, message, RegexOptions.IgnoreCase))
                return true;
            else
                return false;
        }

        public static void PressEnter(long timesToPress = 1)
        {
            for (int i = 0; i < timesToPress; i++)
                Sap.Session.FindById<GuiFrameWindow>("wnd[0]").SendVKey(0); //press enter
        }
    }
}

    