﻿using RpaLib.ProcessAutomation;
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
using RpaLib.Tracing;

namespace RpaLib.SAP
{
    /// <summary>
    /// Class to ease SAP controlling and data scrapping through SAPGui
    /// </summary>
    public class Sap
    {
        private const int _getSapObjTimeoutSeconds = 10;
        private const int _connTimeoutSeconds = 10;
        private const string _defaultFirstTransaction = "SESSION_MANAGER";
        public App App { get; private set; } = new App();
        private Connection _connection;
        public int TriedConnect { get; private set; } = 0;
        public Session FirstSessionAfterConnection { get; private set; }
        public Connection Connection
        {
            get => _connection;
            private set
            {
                _connection = value;
                Trace.WriteLine(string.Join(Environment.NewLine,
                    $"Current connection (Connection) set to:",
                    value));
            }
        }
        public static Connection[] Connections
        {
            get => App.GetConnections();
        }
        public static dynamic SapExe { get; } = new
        {
            BaseName = @"saplogon",
            FullPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        };

        public Sap(Connection connection)
        {
            Connection = connection;
        }

        public Sap(string connection, string transaction)
            : this(connection, user: null, password: null, transaction) { }

        public Sap(string connection, string user, string password, string transaction)
            : this(connection, user, password, client: null, language: null, transaction) { }

        public Sap(string connection, string user, string password, string client, string language, string transaction, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds)
        {
            Connect(connection, connTimeoutSeconds, getSapObjTimeoutSeconds, user, password, client, language);
            FirstSessionAfterConnection = AccessTransaction(transaction);
        }

        #region Connection

        /// <summary>
        /// Low-level method to resolve login after starting a connection.
        /// </summary>
        /// <param name="user">The SAP user's username.</param>
        /// <param name="password">The SAP user's password.</param>
        /// <param name="client">???</param>
        /// <param name="language">???</param>
        private void Login(string user, string password, string client = null, string language = null)
        {
            GuiSession session = Connection.GuiConnection.Children.ElementAt(0) as GuiSession;
            GuiStatusbar statusBar = (GuiStatusbar)session.FindById("wnd[0]/sbar");
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

        /// <summary>
        /// Low-level method to create a new connection to SAP Logon UI using the interop DLL.
        /// </summary>
        /// <param name="connectionName">The connection name (label-like description).</param>
        /// <param name="timeoutSeconds">Timeout in seconds to stay retrying to get the SAP interop object.</param>
        private void CreateSapConnection(string connectionName, int timeoutSeconds = _getSapObjTimeoutSeconds)
        {
            var app = App.GetSapInteropApp(timeoutSeconds);
            Connection = new Connection(app.OpenConnection(connectionName));
            App.Update();
        }
        /// <summary>
        /// Connects to SAP UI object and possibly starts a transaction and do login if parameters were supplied.
        /// </summary>
        /// <param name="connectionName">Label-like connection name description.</param>
        /// <param name="connTimeoutSeconds">Timeout to retry connection by function recall.</param>
        /// <param name="getSapObjTimeoutSeconds">Internal timeout to keep trying to get SAP interop object.</param>
        /// <param name="user">SAP login username.</param>
        /// <param name="password">SAP login password.</param>
        /// <param name="client">??? discover what it does later.</param>
        /// <param name="language">??? discover what it does later.</param>
        /// <exception cref="ExceededRetryLimitSapConnectionException"></exception>
        public void Connect(string connectionName, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds,
            string user = null, string password = null, string client = null, string language = null)
        {
            try
            {
                Trace.WriteLine($"Connecting with SAP: \"{connectionName}\"");
                CreateSapConnection(connectionName, getSapObjTimeoutSeconds);
                Login(user, password, client, language);
                Trace.WriteLine($"Connection Succeeded.\n{Connection}");
            }
            // exception thrown when SAP Connection window is not opened.
            catch (NullReferenceException ex)
            {
                Trace.WriteLine($"Caught exception\n{ex}");
                // retry connection within config's file limit
                TriedConnect++;
                if (TriedConnect <= connTimeoutSeconds)
                {
                    if (TriedConnect == 1)
                    {
                        Trace.WriteLine($"Killing SAP processes by name if it exists: {SapExe.BaseName}");
                        Rpa.KillProcess(SapExe.BaseName);

                        Trace.WriteLine($"Starting SAP exe: {SapExe.FullPath}");
                        Rpa.StartWaitProcess(SapExe.FullPath, outputProcesses: true);
                    }

                    Trace.WriteLine($"Tried connecting to SAP for: {TriedConnect} times. Trying again.");
                    Thread.Sleep(1000);
                    Connect(connectionName, connTimeoutSeconds, getSapObjTimeoutSeconds, user, password, client, language);
                }
                else
                {
                    throw new ExceededRetryLimitSapConnectionException();
                }
            }
        }

        /// <summary>
        /// Search for a GuiConnection in Connections array using its Id as identifier criteria.
        /// </summary>
        /// <param name="connectionId">Connection Id to search for.</param>
        /// <returns></returns>
        public Connection FindConnectionById(string connectionId)
        {
            return App.FindConnectionById(connectionId);
        }

        #endregion

        #region Session

        // high-level function
        public Session CreateNewSession(string connectionName, string transactionId = null, int useSessionId = -1,
            bool iconify = false, bool lockSessionUi = false, int connectionTimeoutSeconds = _connTimeoutSeconds)
        {
            Session session;

            if (Connection == null)
            {
                Connect(connectionName, connectionTimeoutSeconds);
            }
            
            if (useSessionId >= 0)
            {
                // to work with specific session id
                session = Connection.Sessions[useSessionId];
            }
            else
            {
                // to create a new session and work upon it
                session = CreateNewSession();
            }
            
            if (lockSessionUi)
                session.GuiSession.LockSessionUI();
            
            if (iconify)
                session.FindById<GuiFrameWindow>("wnd[0]").Iconify();
            
            if (!string.IsNullOrEmpty(transactionId))
                session.AccessTransaction(transactionId);

            return session;
        }

        // low-level function
        private Session CreateNewSession()
        {
            int lastSession = Connection.Sessions.Length - 1;
            Connection.Sessions[lastSession].CreateNewSession();
            App.Update();
            return FindSession(_defaultFirstTransaction);
        }

        public Session FindSession(string transaction) => Connection.Sessions.Where(x => x.GuiSession.Info.Transaction.Equals(transaction)).FirstOrDefault();

        #endregion

        #region Transaction

        public Session AccessTransaction(string id, string fromTransaction) => AccessTransaction(FindSession(fromTransaction), id);
        public Session AccessTransaction(string id) => AccessTransaction(FindSession(_defaultFirstTransaction), id);
        public static Session AccessTransaction(Session session, string id)
        {
            Trace.WriteLine($"Trying to access transaction \"{id}\"");
            session.GuiSession.StartTransaction(id);
            Trace.WriteLine($"Transaction Access: Successful. Transaction assigned to session: [{session.Index}]");

            return session;
        }

        #endregion

        #region Info

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

        #endregion

        #region UI_Elements

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
                    Trace.WriteLine($"Converting to {typeof(T)} id: {d["PathId"]}");
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

        #endregion

        public void PressEnter(string transaction, long timesToPress = 1)
        {
            Session session = FindSession(transaction);
            for (int i = 0; i < timesToPress; i++)
                session.FindById<GuiFrameWindow>("wnd[0]").SendVKey(0); //press enter
        }
    }
}
