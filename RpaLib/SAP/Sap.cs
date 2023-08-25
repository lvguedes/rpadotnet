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
    public enum TableAction
    {
        Select,
        Click
    }
    public interface ISapTabular
    {
        DataTable DataTable { get; }
        void Parse();
        bool IsEmpty();
        int FulfilledRowsCount();
        string Info();
    }

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

    public class Session
    {
        private int _index;
        public int Index { get => _index; }
        public GuiSession GuiSession { get; set; }

        public Session(GuiSession session)
        {
            GuiSession = session;
            _index = session.Info.SessionNumber;
        }

        public void AccessTransaction(string transactionId) => Sap.AccessTransaction(this, transactionId);
        public T FindById<T>(string id) => (T)FindById(id);

        public T[] FindByText<T>(string labelText) => Sap.FindByText<T>(GuiSession, labelText);

        public void CreateNewSession() => CreateNewSession(this);
        public void CreateNewSession(Session session)
        {
            Tracing.Log.Write($"Creating a new session from Session[{session.Index}]...");
            session.GuiSession.CreateSession();
            Thread.Sleep(2000); // wait otherwise new session cannot be captured
            //Log.MessageBox("New session created.");
            Sap.UpdateConnections();
            Sap.MapExistingSessions();
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"The connection after creating a new session and updating connection through interop engine:",
                Sap.ConnectionInfo()));
            //Log.MessageBox("Verify info");
        }

        public static string SessionInfo(GuiSession session)
        {
            if (session == null) return "No Session";
            return string.Join(Environment.NewLine,
                $"    ApplicationServer: \"{session.Info.ApplicationServer}\"",
                $"    Client: \"{session.Info.Client}\"",
                $"    Codepage: \"{session.Info.Codepage}\"",
                $"    Group: \"{session.Info.Group}\"",
                $"    GuiCodepage: \"{session.Info.GuiCodepage}\"",
                $"    I18NMode: \"{session.Info.I18NMode}\"",
                $"    InterpretationTime: \"{session.Info.InterpretationTime}\"",
                $"    IsLowSpeedConnection: \"{session.Info.IsLowSpeedConnection}\"",
                $"    Language: \"{session.Info.Language}\"",
                $"    MessageServer: \"{session.Info.MessageServer}\"",
                $"    Program: \"{session.Info.Program}\"",
                $"    ResponseTime: \"{session.Info.ResponseTime}\"",
                $"    RoundTrips: \"{session.Info.RoundTrips}\"",
                $"    ScreenNumber: \"{session.Info.ScreenNumber}\"",
                $"    ScriptingModeReadOnly: \"{session.Info.ScriptingModeReadOnly}\"",
                $"    ScriptingModeRecordingDisabled: \"{session.Info.ScriptingModeRecordingDisabled}\"",
                $"    SessionNumber: \"{session.Info.SessionNumber}\"",
                $"    SystemName: \"{session.Info.SystemName}\"",
                $"    SystemNumber: \"{session.Info.SystemNumber}\"",
                $"    SystemSessionId: \"{session.Info.SystemSessionId}\"",
                $"    Transaction: \"{session.Info.Transaction}\"",
                $"    User: \"{session.Info.User}\""
                );
        }

        public override string ToString()
        {
            return
                string.Join(Environment.NewLine,
                $"  Session [{Index}]:",
                SessionInfo(GuiSession)
                );
        }

        public GuiComponent FindById(string pathId)
        {
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Trying to find by id ({pathId})",
                $"The current working session is:",
                this));
            return GuiSession.FindById(pathId);
        }

        public string AllSessionIdsInfo() => Sap.AllSessionIdsInfo(GuiSession);
    }

    public abstract class SapComponent<T> //: Sap
    {
        private string _fullPathId;
        public string Name { get; set; }
        public string Description { get; set; }
        public string FullPathId
        {
            get => _fullPathId;
            set
            {
                Tracing.Log.Write(string.Join(Environment.NewLine,
                    $"Setting the FullPathID of {this.GetType()} \"{Name}\":",
                    $"  Old value: \"{_fullPathId}\"",
                    $"  New value: \"{value}\""));
                _fullPathId = value;
            }
        }
        public string Id { get; set; }
        public static string BasePathId { get; set; }

        public SapComponent()
        {
            //UpdateParentParams();
        }

        public void UpdateParentParams()
        {
            Sap.UpdateConnections();
            Sap.MapExistingSessions();
        }

        public static dynamic FindById(string fullPathId) => Sap.Session.FindById(fullPathId);
        public static U FindById<U>(string fullPathId) => (U)Sap.Session.FindById(fullPathId);
    }

    public class Tab : SapComponent<GuiTab>
    {
        public int Index { get; set; }
        public Field[] Fields { get; set; }

        public void Select() => Select(FullPathId);
        public void Select(string fullPathId)
        {
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Trying to select tab \"{Name}\"",
                $"({fullPathId})"));
            (Sap.Session.FindById(fullPathId) as GuiTab).Select();
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Tab \"{Name}\" selected.",
                $"{fullPathId}"));
        }
    }

    public class Field : SapComponent<GuiTextField>
    {
        public string Xml { get; set; }
        public string Sap { get; set; }
        public Type Datatype { get; set; }
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        public dynamic ConvertToDatatype(string value)
        {
            if (Datatype.Equals(typeof(string)))
                return value;
            else if (Datatype.Equals(typeof(double)))
                return double.Parse(value, new CultureInfo("pt-BR"));
            else if (Datatype.Equals(typeof(long)))
                return long.Parse(value);
            else if (Datatype.Equals(typeof(int)))
                return int.Parse(value);
            else if (Datatype.Equals(typeof(DateTime)))
                return DateTime.ParseExact(value, "dd.MM.yyyy", null);
            else
                throw new FieldDatatypeConversionException();
        }

        public string GetText() => GetText(FullPathId);
        public string GetText(string fullPathId)
        {
            Tracing.Log.Write($"Starting to try to extract field {Name} text. ({FullPathId})");

            string text = (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).Text;

            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Captured field text from some Tab:",
                $"    Field: {Name}",
                $"    Value: {text}"));

            return text;
        }

        public void SetText(string value) => SetText(FullPathId, value);
        public void SetText(string fullPathId, string value)
        {
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).Text = value;
            Tracing.Log.Write($"Field \"{Name}\" text changed. New value: \"{value}\"");
        }

        public void Focus() => Focus(FullPathId);
        public void Focus(string fullPathId)
        {
            Tracing.Log.Write($"Trying to move focus to field \"{Name}\"");
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).SetFocus();
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiTextField).CaretPosition = 0;
            Tracing.Log.Write($"Focus set to field: \"{Name}\". Carret position: 0.");
        }
    }

    public class Button : SapComponent<GuiButton>
    {
        public void Click()
        {
            Tracing.Log.Write($"Trying to click button {Name} ({Description}).");
            Click(FullPathId);
            Tracing.Log.Write($"Button \"{Name}\" ({Description}) clicked.");
        }
        public static void Click(string fullPathId) => (Sap.Session.FindById(fullPathId) as GuiButton).Press();
    }

    public class Table : SapComponent<GuiTableControl>, ISapTabular
    {
        public GuiTableControl GuiTableControl { get; private set; }

        private DataTable dt;
        public DataTable DataTable { get
            {
                //Parse();
                return dt;
            }
            private set
            {
                dt = value;
            }
        }

        public Table()
            : this(name:string.Empty, fullPathId:string.Empty)
        { }

        public Table(string fullPathId) 
            : this(name:string.Empty, fullPathId)
        { }

        public Table(string name, string fullPathId)
        {
            Name = name;
            FullPathId = fullPathId;
            DataTable = new DataTable();
            RefreshTableObj();
            Parse();
            Log.Write(Info());
        }

        private void RefreshTableObj()
        {
            Log.Write("Refreshing the GuiTableControl object...");
            GuiTableControl = FindById<GuiTableControl>(FullPathId);
        }
        private void ResetScrolling()
        {
            Log.Write("Resetting scrollings' position to 0...");
            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.VerticalScrollbar.Position = 0;
                },
                () => { /* Do nothing */ }
            );

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.HorizontalScrollbar.Position = 0;
                },
                () => { /* Do nothing */ }
            );
        }

        public string Info() => Info(this);
        public static string Info(Table table)
        {
            //table.Parse();
            return
                string.Join(Environment.NewLine,
                    $"Table \"{table.Name}\" captured:",
                    $"  CharHeight: \"{table.GuiTableControl.CharHeight}\"",
                    $"  CharWidth:  \"{table.GuiTableControl.CharWidth}\"",
                    $"  CharTop:    \"{table.GuiTableControl.CharTop}\"",
                    $"  CurrentCol: \"{table.GuiTableControl.CurrentCol}\"",
                    $"  CurrentRow: \"{table.GuiTableControl.CurrentRow}\"",
                    $"  RowCount:   \"{table.GuiTableControl.RowCount}\"",
                    $"  Rows.Count: \"{table.GuiTableControl.Rows.Count}\"",
                    $"  VisibleRowCount: \"{table.GuiTableControl.VisibleRowCount}\"",
                    $"  HorizontalScrollbar:",
                    $"    Minimum:  \"{table.GuiTableControl.HorizontalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.HorizontalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.HorizontalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.HorizontalScrollbar.PageSize}\"",
                    $"  VerticalScrollBar:",
                    $"    Minimum:  \"{table.GuiTableControl.VerticalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.VerticalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.VerticalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.VerticalScrollbar.PageSize}\"",
                    $" --- Table object methods ---"
                    ,$"Filled Rows: {table.dt.Rows.Count}"
                    //,$"  FulfilledRowsCount(): {table.FulfilledRowsCount()}"
                    //,$"  IsEmpty(): {table.IsEmpty()}"
                    );
        }

        public static string RowInfo(GuiTableRow row)
        {
            return string.Join(Environment.NewLine,
                $"The first row from table above was captured:)",
                $"  Count:      \"{row.Count}\" (number of cells in the row)",
                $"  Type:       \"{row.Type}\"",
                $"  Selectable: \"{row.Selectable}\" (if the row can be selected)",
                $"  Selected:   \"{row.Selected}\" (if the row is selected)");
        }

        private void ScrollDown(int row)
        {
            int oldPos = GuiTableControl.VerticalScrollbar.Position;
            // relaxed scrolling
            TryEmptyCellActionRelax(
                () =>
                {
                    //rpa.MessageBox("Checking if scroll is needed");
                    Log.Write($"Checking if scroll down is needed for row {row}");
                    if (row % GuiTableControl.VisibleRowCount == 0 && row != 0)
                    {
                        //rpa.MessageBox("Will scroll down now");
                        Log.Write("Will scroll down now...");
                        
                        GuiTableControl.VerticalScrollbar.Position += GuiTableControl.VisibleRowCount;
                        RefreshTableObj();

                        Log.Write(string.Join(Environment.NewLine,
                            $"Scrolling down...",
                            $"  Old position: {oldPos}",
                            $"  New position: {GuiTableControl.VerticalScrollbar.Position}"));
                        //$"  New position: {newPos}"));
                    }
                });

            RefreshTableObj(); // for the case it cannot change position
        }

        private void TryEmptyCellActionRelax(Action tryBlock)
        {
            TryEmptyCellActionRelax(tryBlock, () => { });
        }
        private void TryEmptyCellActionRelax(Action tryBlock, Action catchBlock)
        {
            try
            {
                tryBlock();
            }
            catch (COMException ex)
            {
                if (ex.Message.Equals("The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))"))
                {
                    RefreshTableObj();
                    catchBlock();
                }
                else
                {
                    Log.Write("Unknown COMException occurred." +
                        " It's different from the exception that signals empty cells." +
                        $" See:\n{ex}");
                    throw ex;
                }
            }
        }

        private Dictionary<string, int> GetTableCounters()
        {
            Dictionary<string, int> counters = new Dictionary<string, int>();
            // get columns count or zero if table is empty

            TryEmptyCellActionRelax(
                () =>
                {
                    counters["columns"] = GuiTableControl.Columns.Count;
                },
                () =>
                {
                    counters["columns"] = 0;
                });

            TryEmptyCellActionRelax(
                () =>
                {
                    counters["rows"] = GuiTableControl.RowCount;
                },
                () =>
                {
                    counters["rows"] = 0;
                });

            return counters;
        }

        public GuiVComponent GetCell(int row, int column)
        {
            // The native index is always relative to the number of visible rows 
            // in the screen
            // This function makes it possible to use the real index from any collection
            // to get the cell
            return GuiTableControl.GetCell(row % GuiTableControl.VisibleRowCount, column);
        }

        public void MakeRowVisible(int row)
        {
            int currentPage = GuiTableControl.VerticalScrollbar.Position / GuiTableControl.VisibleRowCount;

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.VerticalScrollbar.Position = row;
                });
            RefreshTableObj();
        }

        public GuiTableRow GetRow(int row)
        {
            MakeRowVisible(row);
            //Rpa.MessageBox($"Will get row {row} now...");
            //return GuiTableControl.Rows.ElementAt(row % GuiTableControl.VisibleRowCount);
            return GuiTableControl.Rows.ElementAt(0);
        }

        public int FulfilledRowsCount()
        {
            int fulfilledRows = 0;
            //Dictionary<string, int> counters = GetTableCounters();
            
            for (int row = 0; row < GetTableCounters()["rows"]; row++)
            {
                //rpa.MessageBox($"Processing line {row}");
                ScrollDown(row);

                for(int col = 0; col < GetTableCounters()["columns"]; col++)
                {
                    Log.Write($"Iteration: row {row}, col {col}");
                    bool isCellNullOrEmpty = false;
                    TryEmptyCellActionRelax(
                        () =>
                        {
                            if ( !string.IsNullOrEmpty(
                                     GetCell(row, col).Text) )
                            {
                                fulfilledRows++;
                                isCellNullOrEmpty = true;
                            }
                        },
                        () =>
                        {
                            isCellNullOrEmpty = true;
                        });
                    if (isCellNullOrEmpty) break;
                    /*
                    try
                    {
                        if (!string.IsNullOrEmpty(GuiTableControl.GetCell(row, col).Text))
                        {
                            fulfilledRows++;
                            break;
                        }
                    }
                    catch(COMException ex)
                    {
                        if (ex.Message.Equals("The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))"))
                        {
                            Log.Write("Caught a null empty cell. Skipping...");
                            RefreshTableObj();
                            break;
                        }
                        else
                        {
                            Log.Write("Unknown COMException occurred." +
                                " It's different from the exception that signals empty cells." +
                                $" See:\n{ex}");
                            throw ex;
                        }
                    }
                    */
                }

                // if last line is empty, stop analyzing lines
                if (fulfilledRows != row + 1)
                    break;
            }

            ResetScrolling();

            return fulfilledRows;
        }

        public bool IsEmpty()
        {
            bool isEmpty = false;

            if (FulfilledRowsCount() == 0) 
                isEmpty = true;

            Log.Write($"Table \"{Name}\" is empty? {isEmpty}");

            return isEmpty;
        }

        public void Parse()
        {
            RefreshTableObj();
            dt = new DataTable();

            foreach (GuiTableColumn col in GuiTableControl.Columns)
            {
                dt.Columns.Add(col.Title, typeof(string));
            }

            int fulfilledRowsCount = FulfilledRowsCount();
            for (int row = 0; row < fulfilledRowsCount; row++)
            {
                ScrollDown(row);

                DataRow datarow = dt.NewRow();
                for (int col = 0; col < GetTableCounters()["columns"]; col++)
                {
                    //rpa.MessageBox($"Getting cell ({row}, {col})...");
                    datarow[col] = GetCell(row, col).Text;
                }
                dt.Rows.Add(datarow);
            }
            ResetScrolling();
        }

        public string PrintDataTable() => Rpa.PrintDataTable(dt);
/*
        public void Find(TableAction action)
        {


            for (long i = 0; i < GuiTableControl.RowCount; i++)
                for (long j = 0; j < GuiTableControl.Columns.Count; j++)
                    if (action == TableAction.Select) 
        }
*/
    }


    public class Grid : SapComponent<GuiGridView>, ISapTabular
    {
        public GuiGridView GuiGridView { get; private set; }
        public DataTable DataTable { get; private set; }

        public Grid() { }

        public Grid(string fullPathId) : this()
        {
            FullPathId = fullPathId;
            Refresh();
        }

        // refreshes datatable info after updating the GuiGridView object
        public void Refresh()
        {
            DataTable = new DataTable();
            GuiGridView = FindById(FullPathId);
            Parse();
            Log.Write(Info());
        }

        public void Parse()
        {
            GuiGridView.SelectAll();
            //Cols = _grid.SelectedColumns;
            //Rows = long.Parse(Regex.Match(_grid.SelectedRows, @"(?<=\d+-)\d+").Value) + 1;

            foreach (var col in GuiGridView.SelectedColumns)
            {
                DataTable.Columns.Add(col, typeof(string));
            }

            for (int i = 0; i < GuiGridView.RowCount; i++)
            {
                DataRow datarow = DataTable.NewRow();

                if (i != 0 && i % GuiGridView.VisibleRowCount == 0)
                {
                    GuiGridView.FirstVisibleRow = i;
                }

                foreach (DataColumn col in DataTable.Columns)
                {
                    //_grid.SetCurrentCell(i, col);
                    datarow[col.ColumnName] = GuiGridView.GetCellValue(i, col.ColumnName);
                }
                DataTable.Rows.Add(datarow);
            }
        }

        public bool IsEmpty()
        {
            if (DataTable.Rows.Count == 0)
                return true;
            else
                return false;
        }

        public int FulfilledRowsCount() => DataTable.Rows.Count;


        public string Info() => Info(this);
        public static string Info(Grid grid)
        {
            string[] columnNamesDatatable = grid.DataTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
            string[] selectedColumns = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.SelectedColumns).ToArray();
            string[] selectedCells = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.SelectedCells).ToArray();
            string[] columnOrder = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.ColumnOrder).ToArray();

            return
              string.Join(Environment.NewLine,
                  $"Grid \"{grid.Name}\" captured:",
                  $"  ColumnCount:         \"{grid.GuiGridView.ColumnCount}\"",
                  $"  ColumnOrder :        \"{string.Join(", ", columnOrder)}\"",
                  $"  CurrentCellColumn :  \"{grid.GuiGridView.CurrentCellColumn}\"",
                  $"  CurrentCellRow :     \"{grid.GuiGridView.CurrentCellRow}\"",
                  $"  FirstVisibleColumn : \"{grid.GuiGridView.FirstVisibleColumn}\"",
                  $"  FirstVisibleRow :    \"{grid.GuiGridView.FirstVisibleRow}\"",
                  $"  FrozenColumnCount :  \"{grid.GuiGridView.FrozenColumnCount}\"",
                  $"  RowCount:            \"{grid.GuiGridView.RowCount}\"",
                  $"  SelectedCells:       \"{string.Join(", ", selectedCells)}\"",
                  $"  SelectedColumns:     \"{string.Join(", ", selectedColumns)}\"",
                  $"  SelectedRows:        \"{grid.GuiGridView.SelectedRows}\"",
                  $"  SelectionMode:       \"{grid.GuiGridView.SelectionMode}\"",
                  $"  Title:               \"{grid.GuiGridView.Title}\"",
                  $"  ToolbarButtonCount:  \"{grid.GuiGridView.ToolbarButtonCount}\"",
                  $"  VisibleRowCount:     \"{grid.GuiGridView.VisibleRowCount}\"",
                  $"  Grid.DataTable.Columns.Count: {string.Join(", ", columnNamesDatatable)}",
                  $"  Grid.DataTable.Rows.Count: \"{grid.DataTable.Rows.Count}\""
                  );
        }

        public string PrintDataTable() => Rpa.PrintDataTable(DataTable);

        /*
        TODO: 
        Methods to check:
        string GetCellState(long Row, string Column) {Possible return: Normal, Error, Warning, Info}
        string GetCellType(long Row, string Column) {Possible return: Normal, Button, Checkbox, ValueList, RadioButton}
        string GetCellValue(long Row, string Column)
        ICollection<string> GetColumnTitles(string Column) // all possibilities of given column titles
        string GetDisplayedColumnTitle(string Column) // one of the values returned by function above
        void SetCurrentCell(long Row, string Column)

        void SelectAll() // selects all columns and rows
        ICollection<string> SelectedColumns { get; set; } // Setting this property can raise an exception, if the new collection contains an invalid column identifier.

        Properties:
        long CurrentCellRow { get; set; }
        string CurrentCellColumn { get; set; }
         */
    }
}

    