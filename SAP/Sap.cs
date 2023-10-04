using RpaLib.ProcessAutomation;
using sapfewse;
using saprotwr.net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using SysTrace = System.Diagnostics.Trace;

namespace RpaLib.SAP
{
    public class Sap
    {
        private const int _getSapObjTimeoutSeconds = 10;
        private const int _connTimeoutSeconds = 10;
        private GuiConnection _connection;
        public Session[] Sessions { get; set; }
        public int TriedConnect { get; private set; } = 0;
        public GuiApplication App { get; set; }
        public GuiConnection Connection
        {
            get => _connection;
            set
            {
                _connection = value;
                SysTrace.WriteLine(string.Join(Environment.NewLine,
                    $"Current connection (Connection) set to:",
                    ConnectionInfo(value, this)));
            }
        }
        public static GuiConnection[] Connections { get; set; }
        public static dynamic SapExe { get; } = new
        {
            BaseName = @"saplogon",
            FullPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        };

        public Sap(string connection, string transaction)
            : this(connection, user: null, password: null, transaction) { }

        public Sap(string connection, string user, string password, string transaction)
            : this(connection, user, password, client: null, language: null, transaction) { }

        public Sap(string connection, string user, string password, string client, string language, string transaction, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds)
        {
            Connect(connection, connTimeoutSeconds, getSapObjTimeoutSeconds, user, password, client, language);
            AccessTransaction(transaction);
        }

        private static GuiApplication GetSapInteropApp(int timeout = _getSapObjTimeoutSeconds)
        {
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object sapGuilRot = null;
            int timePassed = 0;
            
            // try to get SapGuilRot within timeout
            while(sapGuilRot == null && timePassed < timeout)
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

        private void Login(string user, string password, string client = null, string language = null)
        {
            GuiSession session = Connection.Children.ElementAt(0) as GuiSession;
            GuiStatusbar statusBar = (GuiStatusbar)session.FindById("/app/con[0]/ses[0]/wnd[0]/sbar");
            GuiMainWindow mainWindow = (GuiMainWindow)session.FindById("wnd[0]");

            if (client != null)
                (session.FindById("wnd[0]/usr/txtRSYST-MANDT") as GuiVComponent).Text = client;
            if (user != null)
                (session.FindById("wnd[0]/usr/txtRSYST-BNAME") as GuiVComponent).Text = user;
            if (password != null)
                (session.FindById("wnd[0]/usr/pwdRSYST-BCODE") as GuiVComponent).Text = password;
            if (language != null)
                (session.FindById("wnd[0]/usr/txtRSYST-LANGU") as GuiVComponent).Text = language;

            if (client != null || user != null || password != null || language != null)
                mainWindow.SendVKey(0); //press Enter
        }

        private void CreateSapConnection(string connectionName, int timeoutSeconds = _getSapObjTimeoutSeconds)
        {
            App = GetSapInteropApp(timeoutSeconds);

            if (App.Connections.Count == 0)
            {
                Connection = App.OpenConnection(connectionName);
            }
            else
            {
                Connection = App.Connections.ElementAt(0) as GuiConnection;
            }
        }

        public void Connect(string connectionName, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds,
            string user = null, string password = null, string client = null, string language = null)
        {
            try
            {
                SysTrace.WriteLine($"Connecting with SAP: \"{connectionName}\"");
                CreateSapConnection(connectionName, getSapObjTimeoutSeconds);
                Login(user, password, client, language);
                MapExistingSessions();
                SysTrace.WriteLine($"Connection Succeeded.\n{ConnectionInfo()}");
            }
            // exception thrown when SAP Connection window is not opened.
            catch (NullReferenceException ex)
            {
                SysTrace.WriteLine($"Caught exception\n{ex}");
                // retry connection within config's file limit
                TriedConnect++;
                if (TriedConnect <= connTimeoutSeconds)
                {
                    if (TriedConnect == 1)
                    {
                        SysTrace.WriteLine($"Killing SAP processes by name if it exists: {SapExe.BaseName}");
                        Rpa.KillProcess(SapExe.BaseName);

                        SysTrace.WriteLine($"Starting SAP exe: {SapExe.FullPath}");
                        Rpa.StartWaitProcess(SapExe.FullPath, outputProcesses: true);
                    }

                    SysTrace.WriteLine($"Tried connecting to SAP for: {TriedConnect} times. Trying again.");
                    Thread.Sleep(1000);
                    Connect(connectionName, connTimeoutSeconds, getSapObjTimeoutSeconds, user, password, client, language);
                }
                else
                {
                    throw new ExceededRetryLimitSapConnectionException();
                }
            }
        }

        // create a new connection or uses the current actually? (if it creates this is an error)
        public void UpdateConnections()
        {
            List<GuiConnection> connections = new List<GuiConnection>();
            App = GetSapInteropApp();

            foreach (GuiConnection conn in App.Connections)
            {
                connections.Add(conn);
            }
            Connections = connections.ToArray();

            SysTrace.WriteLine(string.Join(Environment.NewLine,
                $"Updating connections...",
                $"The number of existing connections: {connections.Count}",
                $"Available connections:",
                ConnectionsInfo(this)));

            Connection = App.Connections.ElementAt(0) as GuiConnection;

        }

        public void MapExistingSessions()
        {
            List<Session> sessions = new List<Session>();

            SysTrace.WriteLine($"Trying to map existing sessions. Number of existing sessions: {Connection.Sessions.Count}");

            foreach (GuiSession session in Connection.Sessions)
            {
                SysTrace.WriteLine(string.Join(Environment.NewLine,
                    "Mapping session:",
                    Session.SessionInfo(session)));
                sessions.Add(new Session(session, this));
            }

            // sort by Session.GuiSession.Info.SessionNumber
            Sessions = sessions.OrderBy(x => x.GuiSession.Info.SessionNumber).ToArray();
        }

        // high-level function
        public Session CreateNewSession(string connectionName, string transactionId = null, int useSessionId = -1, int connectionTimeoutSeconds = _connTimeoutSeconds)
        {
            Session session;

            if (Connection == null)
            {
                Connect(connectionName, connectionTimeoutSeconds);
            }
            if (useSessionId >= 0)
            {
                // to work with specific session id
                session = Sessions[0];
            }
            else
            {
                // to create a new session and work upon it
                session = CreateNewSession();
            }
            //session.GuiSession.LockSessionUI();
            session.FindById<GuiFrameWindow>("wnd[0]").Iconify();
            if (!string.IsNullOrEmpty(transactionId))
                session.AccessTransaction(transactionId);

            return session;
        }

        // low-level function
        private Session CreateNewSession()
        {
            int lastSession = Sessions.Length - 1;
            Sessions[lastSession].CreateNewSession();
            return FindSession("SESSION_MANAGER");
        }

        public Session FindSession(string transaction) => Sessions.Where(x => x.GuiSession.Info.Transaction.Equals(transaction)).FirstOrDefault();

        public void AccessTransaction(string id, string fromTransaction) => AccessTransaction(FindSession(fromTransaction), id);
        public void AccessTransaction(string id) => AccessTransaction(FindSession("SESSION_MANAGER"), id);
        public void AccessTransaction(Session session, string id)
        {
            SysTrace.WriteLine($"Trying to access transaction \"{id}\"");
            session.GuiSession.StartTransaction(id);
            SysTrace.WriteLine($"Transaction Access: Successful. Transaction assigned to session: [{session.Index}]");
        }

        public string ConnectionInfo() => ConnectionInfo(Connection, this);
        public static string ConnectionInfo(GuiConnection connection, Sap sap2obj)
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
                sap2obj.SessionsInfo()
                );
        }

        public static string ConnectionsInfo(Sap sap2obj)
        {
            StringBuilder info = new StringBuilder();

            for (int i = 0; i < Connections.Length; i++)
            {
                info.AppendLine(string.Join(Environment.NewLine,
                    $"[{i}]",
                    ConnectionInfo(Connections[i], sap2obj)
                    ));
            }
            return info.ToString();
        }

        public string SessionsInfo()
        {
            StringBuilder info = new StringBuilder();
            if (Sessions == null) return "No Sessions";
            foreach (Session session in Sessions)
            {
                info.AppendLine(string.Concat(Environment.NewLine, session));
            }

            return info.ToString();
        }

        public string MainWindowInfo(string transaction)
        {
            Session session = FindSession(transaction);
            GuiMainWindow mw = session.FindById("wnd[0]") as GuiMainWindow;
            return string.Join(Environment.NewLine,
                $"GuiMainWindow info:",
                $"  ButtonbarVisible: {mw.ButtonbarVisible}",
                $"  StatusbarVisible: {mw.StatusbarVisible}",
                $"  TitlebarVisible:  {mw.TitlebarVisible}",
                $"  ToolbarVisible:   {mw.ToolbarVisible}"
                );
        }

        public string AllSessionIdsInfo(string transaction)
        {
            return AllSessionIdsInfo(FindSession(transaction).GuiSession);
        }

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
                .Select((d) => {
                    SysTrace.WriteLine($"Converting to {typeof(T)} id: {d["PathId"]}");
                    try
                    {
                        return (T)d["Obj"];
                    }
                    catch (InvalidCastException)
                    {
                        return default(T);
                    }
                })
                .Where(x => x != null)
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
            List<Dictionary<string, dynamic>> ids = new List<Dictionary<string, dynamic>>();

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

        public string CurrentStatusBarInfo(string transaction)
        {
            Session session = FindSession(transaction);
            GuiStatusbar currentStatusBar = session.FindById<GuiStatusbar>("wnd[0]/sbar");

            return string.Join(Environment.NewLine,
                $"Current Status Bar properties:",
                $"  MessageType: {currentStatusBar.MessageType}",
                $"  Text: {currentStatusBar.Text}");
        }
        public bool IsStatusType(StatusType status, string transaction)
        {
            Session session = FindSession(transaction);
            string statusLetter = StatusTypeEnum.GetStatusTypeLetter(status);
            if (Regex.IsMatch(session.FindById<GuiStatusbar>("wnd[0]/sbar").MessageType, statusLetter, RegexOptions.IgnoreCase))
                return true;
            else
                return false;
        }

        public bool IsStatusMessage(string message, string transaction)
        {
            Session session = FindSession(transaction);
            if (Regex.IsMatch(session.FindById<GuiStatusbar>("wnd[0]/sbar").Text, message, RegexOptions.IgnoreCase))
                return true;
            else
                return false;
        }

        public void PressEnter(string transaction, long timesToPress = 1)
        {
            Session session = FindSession(transaction);
            for (int i = 0; i < timesToPress; i++)
                session.FindById<GuiFrameWindow>("wnd[0]").SendVKey(0); //press enter
        }
    }
}
