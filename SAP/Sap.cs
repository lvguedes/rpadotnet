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
using RpaLib.Tracing;
using RpaLib.SAP.Model;
using RpaLib.SAP.Exceptions;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RpaLib.SAP
{
    /// <summary>
    /// Class to ease SAP controlling and data scrapping through SAPGui.
    /// This class is based in a single connection that must be passed somehow to the constructor when instantiating.
    /// You must create a object of this class for each connection you want to manage.
    /// </summary>
    public class Sap
    {
        private const int _getSapObjTimeoutSeconds = 10;
        private const int _connTimeoutSeconds = 10;
        private const string _defaultFirstTransaction = "SESSION_MANAGER";

        /// <summary>
        /// Manage the Application COM object through the wrapper class App.
        /// </summary>
        public App App { get; private set; } = new App();

        /// <summary>
        /// Backing field for the GuiConnection wrapper.
        /// </summary>
        private Connection _connection;

        /// <summary>
        /// How many times a connection was tried to be established.
        /// </summary>
        public int TriedConnect { get; private set; } = 0;

        /// <summary>
        /// The first session right after a new connection is opened.
        /// </summary>
        public Session FirstSessionAfterConnection { get; private set; }

        /// <summary>
        /// Connection name (description) from the GuiConnection object this class manages.
        /// The connection name is the same that appears in the connection description in GUI.
        /// </summary>
        public string ConnectionName { get => Connection.GuiConnection.Description; }

        /// <summary>
        /// Sap connection managed by this class. All other methods will refer to this connection.
        /// </summary>
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

        /// <summary>
        /// All connections defined in the GuiApplication object got through its wrapper class App.
        /// </summary>
        public static Connection[] Connections
        {
            get => App.GetConnections();
        }

        /// <summary>
        /// Inner class that model SAP executable info
        /// </summary>
        public class SapExeInfo
        {
            public string BaseName { get; set; }
            public string FullPath { get; set; }
        }

        /// <summary>
        /// Property to get SAP executable info.
        /// </summary>
        public static SapExeInfo SapExe { get; } = new SapExeInfo
        {
            BaseName = @"saplogon",
            FullPath = @"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        };

        /// <summary>
        /// Initiate the object of this class with the given connection.
        /// </summary>
        /// <param name="connection"></param>
        public Sap(Connection connection)
        {
            Connection = connection;
        }

        /// <summary>
        /// Open a new connection with connection description and transaction name. Don't make login, SAP user must be signed in.
        /// </summary>
        /// <param name="connectionName">The description name of the connection. The same that appear in GUI.</param>
        /// <param name="transaction">The transaction name that should be accessed after opening the connection.</param>
        public Sap(string connectionName, string transaction)
            : this(connectionName, user: null, password: null, transaction) { }

        /// <summary>
        /// Open a new SAP connection and login to SAP.
        /// </summary>
        /// <param name="connectionName">The connection description that appear in GUI.</param>
        /// <param name="user">The SAP user name.</param>
        /// <param name="password">The SAP user password.</param>
        /// <param name="transaction">The transaction name that must be accessed once connected.</param>
        public Sap(string connectionName, string user, string password, string transaction)
            : this(connectionName, user, password, client: null, language: null, transaction) { }

        /// <summary>
        /// Open a new SAP connection and login. Additionally enable you to set up advanced settings to modify this class internal behavior.
        /// </summary>
        /// <param name="connectionName">The connection description that appear in GUI</param>
        /// <param name="user">The SAP user name.</param>
        /// <param name="password">The SAP user password.</param>
        /// <param name="client">???</param>
        /// <param name="language">???</param>
        /// <param name="transaction">The transaction name that must be accessed once connected.</param>
        /// <param name="connTimeoutSeconds">Try to connect by default 10 times waiting 1 second per attempt. 
        ///                                  The attempts per second can be changed through this parameter.</param>
        /// <param name="getSapObjTimeoutSeconds">Seconds to stay retrying to get the SAP object when the process is still starting.</param>
        public Sap(string connectionName, string user, string password, string client, string language, string transaction, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds)
        {
            Connect(connectionName, connTimeoutSeconds, getSapObjTimeoutSeconds, user, password, client, language);
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
        /// <param name="connectionName">Label-like connection name description (the same that appear in the description in GUI).</param>
        /// <param name="connTimeoutSeconds">Timeout to retry connection by function recall.</param>
        /// <param name="getSapObjTimeoutSeconds">Internal timeout to keep trying to get SAP interop object.</param>
        /// <param name="user">SAP login username.</param>
        /// <param name="password">SAP login password.</param>
        /// <param name="client">??? discover what it does later.</param>
        /// <param name="language">??? discover what it does later.</param>
        /// <exception cref="ExceededRetryLimitSapConnectionException"></exception>
        private void Connect(string connectionName, int connTimeoutSeconds = _connTimeoutSeconds, int getSapObjTimeoutSeconds = _getSapObjTimeoutSeconds,
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
        /// <returns>The Connection wrapper object found within App connections collection.</returns>
        public Connection FindConnectionById(string connectionId)
        {
            return App.FindConnectionById(connectionId);
        }

        /// <summary>
        /// Close the connection managed by this class. Close the connection along with all its sessions.
        /// </summary>
        public void CloseConnection() => Connection.GuiConnection.CloseConnection();

        #endregion

        #region Session

        /// <summary>
        /// Create a new session using the last session of Sessions as default creator.
        /// A new session makes a new SAP Frame window to open and a transaction be starte with it.
        /// The first transaction opened is always the "SESSION_MANAGER", but you can also specify
        /// a different transaction to be opened.
        /// </summary>
        /// <param name="transactionId">Transaction to be accessed after the connection is opened.</param>
        /// <param name="useSessionId">Session ID to use as the new session creator.</param>
        /// <param name="iconify">Option to reduce the frame window size to its minimum. Useful when trying to block user interactions when automating.</param>
        /// <param name="lockSessionUi">Option to lock Session User Interface so that no user interaction is possible until the session is unlocked
        ///                             using UnlockSessionUI from GuiSession.</param>
        /// <param name="connectionTimeoutSeconds">Try to connect by default 10 times waiting 1 second per attempt. 
        ///                                        The attempts per second can be changed through this parameter.</param>
        /// <returns></returns>
        public Session CreateNewSession(string transactionId = null, int useSessionId = -1,
            bool iconify = false, bool lockSessionUi = false, int connectionTimeoutSeconds = _connTimeoutSeconds)
        {
            Session session;
            
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

        /// <summary>
        /// Search for the first session which has the transaction.
        /// </summary>
        /// <param name="transaction">Transaction name.</param>
        /// <returns>A Session (wrapper to GuiSession) object.</returns>
        public Session FindSession(string transaction) => Connection.Sessions.Where(x => x.GuiSession.Info.Transaction.Equals(transaction)).FirstOrDefault();

        /// <summary>
        /// A session can be closed by calling this method of the connection. Closing the last session of a connection will close the connection, too.
        /// </summary>
        /// <param name="sessionId">The id of the session to close (like "/app/con[0]/ses[0]")</param>
        public void CloseSession(string sessionId) => Connection.GuiConnection.CloseSession(sessionId);

        #endregion

        #region Transaction

        /// <summary>
        /// Find the session with transaction and access the new transaction
        /// </summary>
        /// <param name="newTransaction">New transaction to be accessed.</param>
        /// <param name="sessionWithTransaction">Session with this transaction will be found and changed to the new transaction.</param>
        /// <returns>A Session wrapper containing the GuiSession object that controls the opened transaction.</returns>
        public Session AccessTransaction(string newTransaction, string sessionWithTransaction) => AccessTransaction(FindSession(sessionWithTransaction), newTransaction);

        /// <summary>
        /// Change the transaction of the session opened after establishing the SAP connection.
        /// </summary>
        /// <param name="transaction">The name of the new transaction to be switched for.</param>
        /// <returns>A Session wrapper containing the GuiSession object.</returns>
        public Session AccessTransaction(string transaction) => AccessTransaction(FindSession(_defaultFirstTransaction), transaction);

        /// <summary>
        /// Access the transaction within the given session.
        /// </summary>
        /// <param name="session">Session in which we need to change the transaction</param>
        /// <param name="transaction">The transaction name of the new transaction the session is going to access.</param>
        /// <returns></returns>
        public static Session AccessTransaction(Session session, string transaction)
        {
            Trace.WriteLine($"Trying to access transaction \"{transaction}\"");
            session.GuiSession.StartTransaction(transaction);
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
            return AllDescendantsInfo(session);
        }

        public static string AllDescendantsInfo(dynamic guiContainer)
        {
            SapGuiObject[] ids = AllDescendants(guiContainer);
            StringBuilder info = new StringBuilder($"All IDs from \"[{guiContainer.Type}]: {guiContainer.Id}\" to its innermost descendant are:\n");

            foreach (var id in ids)
            {
                info.AppendLine($"[{id.Type}]: {id.PathId} \"{id.Text}\"");
            }

            return info.ToString();
        }

        #endregion

        #region UI_Elements

        public static T FindById<T>(GuiComponent parent, string pathId, bool showTypes = false)
        {
            T foundObj;

            if (parent.ContainerType)
                foundObj = parent is GuiContainer ? (T)(parent as GuiContainer).FindById(pathId) : (T)(parent as GuiVContainer);
            else
                throw new ArgumentException($"The argument parent must be a Container type (GuiContainer or GuiVContainer).");

            if (showTypes)
            {
                Trace.WriteLine(string.Join("\n",
                    $"Type of obj found: \"{((GuiComponent)foundObj).Type}\"",
                    $"Is ContainerType? {((GuiComponent)foundObj).ContainerType}",
                    $"Is GuiContainer? {foundObj is GuiContainer}",
                    $"Is GuiVContainer? {foundObj is GuiVContainer}",
                    $"Is GuiComponent? {foundObj is GuiComponent}",
                    $"Is GuiVComponent? {foundObj is GuiVComponent}"), color: ConsoleColor.Yellow);
            }

            return foundObj;
        }

        /// <summary>
        /// Search for an element by matching its text with a regex pattern parameter.
        /// </summary>
        /// <typeparam name="T">The type of the Sap element you're looking for.</typeparam>
        /// <param name="parent">The parent object in which to look for the element.</param>
        /// <param name="labelTextRegex">Regex pattern to search within session. The first found will be returned.</param>
        /// <returns>An array containing the SAP Gui elements found by text.</returns>
        public static T[] FindByText<T>(GuiComponent parent, string labelTextRegex)
        {
            T[] objFound = AllDescendants(parent)
                .Cast<SapGuiObject>()
                .Where(elt =>
                {
                    GuiVComponent visualComponent;
                    try
                    {
                        visualComponent = (GuiVComponent)elt.Obj;
                        return Rpa.IsMatch(visualComponent.Text, labelTextRegex);
                    }
                    catch (InvalidCastException)
                    {
                        // if not a visual component, there isn't text
                        return false;
                    }
                })
                .Select((elt) => {
                    Trace.WriteLine($"Converting [{elt.Type}]{elt.PathId} \"{elt.Text}\" to {typeof(T)}");
                    try
                    {
                        return (T)elt.Obj;
                    }
                    catch (InvalidCastException)
                    {
                        return default(T);
                    }
                })
                .Where(x => x != null)
                .ToArray();

            /* 
            // Not good when quering the result
            if (objFound.Count() == 0)
                objFound = null;
            */

            return objFound;
        }

        public static T[] FindByType<T>(dynamic rootContainer, bool showFound = false) => FindByType<T>(rootContainer, typeof(T).ToString(), showFound);

        public static T[] FindByType<T>(dynamic rootContainer, string typeName, bool showFound = false)
        {
            SapGuiObject[] descendants = AllDescendants(rootContainer);

            var descendantsFound = from descendant in descendants
                                   where Rpa.IsMatch(typeName, descendant.Type)
                                   select descendant;

            if (showFound)
            {
                var descendantTypeRealtype = string.Join(",\n\t", descendantsFound.Select(x => $"{x.Type} ({x.Obj})"));
                var root = (GuiComponent)rootContainer;
                var rootText = rootContainer is GuiVComponent ? (rootContainer as GuiVComponent)?.Text : string.Empty;
                Trace.WriteLine($"Looking for type [{typeName}] within object \"{root.Id}\"...");
                Trace.WriteLine($"[{root.Type}] {root.Id} \"{rootText}\":\n\t{descendantTypeRealtype}");
            }

            return descendantsFound.Select(x => (T)x.Obj).ToArray();
        }

        public static bool ExistsByText<T>(GuiComponent parent, string textRegex)
        {
            var resultList = FindByText<T>(parent, textRegex);
            if (resultList.Length == 0)
                return false;
            else
                return true;
        }

        public static bool ExistsById<T>(GuiComponent parent, string pathId)
        {
            try
            {
                FindById<T>(parent, pathId);
            }
            catch (COMException ex)
            {
                if (Rpa.IsMatch(ex.Message, @"The control could not be found by id\."))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Verifies if a text exists inside any descendant objects of a parent object which is descendant of this Session.GuiSession.
        /// </summary>
        /// <typeparam name="P">Type of the parent object.</typeparam>
        /// <typeparam name="C">Type of the object that contains a Text property in which the regex passed as parameter must match.
        /// Note that it isn't the type of the parent object, instead, it's the type of its child that must contain the regex like text.</typeparam>
        /// <param name="root">A GuiContainer from where to start the search.</param>
        /// <param name="parentPathId">Path ID from the parent object which is inside root. All its children are considered recursively (children of children, etc.)</param>
        /// <param name="textRegex">Pattern which any parent descendant's Text property must match</param>
        /// <returns>Boolean indicating if there is at least one descendant with the text specified.</returns>
        public static bool ExistsTextInside<P, C>(GuiComponent root, string parentPathId, string textRegex)
        {
            if (!ExistsById<P>(root, parentPathId))
                return false;

            try
            {
                var foundObj = FindById<P>(root, parentPathId, showTypes: false);

                return ExistsByText<C>((GuiComponent)foundObj, textRegex);
            }
            catch (InvalidCastException)
            {
                return false;
            }
        }

        public static C[] FindTextInside<P, C>(GuiComponent root, string parentPathId, string textRegex)
        {
            return FindByText<C>((GuiComponent)FindById<P>(root, parentPathId), textRegex);
        }

        /// <summary>
        /// Get all children objects from a given GuiSession recursively.
        /// </summary>
        /// <param name="session">The GuiSession object to get the children objects tree as a list.</param>
        /// <returns>A list of dictionaries containing info about each child object found.</returns>
        public static SapGuiObject[] AllSessionIds(GuiSession session) => AllDescendants(session);

        /// <summary>
        /// Get all children objects recursively from a root object.
        /// </summary>
        /// <param name="rootContainer">The root object to start parsing the children tree.</param>
        /// <returns>A list of dictionaries containing info about each child object found.</returns>
        public static SapGuiObject[] AllDescendants(dynamic rootContainer)
        {
            List<SapGuiObject> objects = new List<SapGuiObject>();

            AllDescendants(objects, rootContainer);

            return objects.ToArray();
        }

        /// <summary>
        /// Auxiliar method to get all children objects recursively from a root object.
        /// </summary>
        /// <param name="objects">List of captured IDs.</param>
        /// <param name="root">The root object that will be changed through each recursive call.</param>
        /// <returns>A list of dictionaries containing info about each child object found.</returns>
        private static List<SapGuiObject> AllDescendants(List<SapGuiObject> objects, dynamic root)
        {
            Action<dynamic> addNodeToList =
                (dynamic node) =>
                {
                    objects.Add(
                        new SapGuiObject
                        {
                            PathId = (string)node.Id,
                            Type = (string)node.Type,
                            Text = node is GuiVComponent? node.Text : string.Empty,
                            Obj = node,
                        });
                };

            for (int i = 0; i < root.Children.Count; i++)
            {
                dynamic child = root.Children.ElementAt[i];
                if (child.ContainerType)
                    AllDescendants(objects, child);
                else
                    addNodeToList(child);
            }

            addNodeToList(root);

            return objects;
        }

        #endregion


        public void PressEnter(string transaction, long timesToPress = 1, int wndIndex = 0)
        {
            Session session = FindSession(transaction);
            for (int i = 0; i < timesToPress; i++)
                session.FindById<GuiFrameWindow>($"wnd[{wndIndex}]").SendVKey(0); //press enter
        }
    }
}
