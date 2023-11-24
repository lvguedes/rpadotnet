using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using RpaLib.Tracing;
using RpaLib.ProcessAutomation;
using RpaLib.SAP.Model;
using System.Runtime.InteropServices;

namespace RpaLib.SAP
{
    public class Session
    {
        private int _index;
        private const int _sessionWaitMilisec = 2000;
        public Connection Connection { get; private set; }
        public int Index { get => _index; }
        public GuiSession GuiSession { get; set; }
        public GuiStatusbar CurrentStatusBar
        {
            get
            {
                return FindById<GuiStatusbar>("wnd[0]/sbar");
            }
        }
        public GuiFrameWindow CurrentFrameWindow
        {
            get
            {
                return GetWindow(0);
            }
        }

        public string Transaction
        {
            get => GuiSession.Info.Transaction;
        }

        public ModalWindow[] PopUps
        {
            get
            {
                List<ModalWindow> modalWindows = new List<ModalWindow>();
                
                var foundPopUps = FindByType<GuiModalWindow>(showFound: true);

                foreach (var popUp in foundPopUps)
                {
                    modalWindows.Add(new ModalWindow(popUp.Id, this));
                }

                return modalWindows.ToArray();
            }
        }

        public Session(GuiSession guiSession, Connection connection)
        {
            GuiSession = guiSession;
            _index = guiSession.Info.SessionNumber;
            Connection = connection;
        }

        public void SendVKey(int vKeyNumber) => CurrentFrameWindow.SendVKey(vKeyNumber);

        public void AccessTransaction(string transactionId) => Sap.AccessTransaction(this, transactionId);

        public GuiComponent FindById(string pathId, bool suppressTrace = false)
        {
            if (! suppressTrace)
                Trace.WriteLine(string.Join(Environment.NewLine,
                    $"Trying to find by id ({pathId})",
                    $"The current working session is:",
                    this));
            return GuiSession.FindById(pathId);
        }

        public T FindById<T>(string id, bool suppressTrace = false) => (T)FindById(id, suppressTrace);

        /// <summary>
        /// Search for an element by matching its text with a regex pattern parameter.
        /// </summary>
        /// <typeparam name="T">The type of the Sap element you're looking for.</typeparam>
        /// <param name="labelTextPattern">Regex pattern to search within session. The first found will be returned.</param>
        /// <returns></returns>
        public T[] FindByText<T>(string labelTextPattern) => Sap.FindByText<T>((GuiComponent)GuiSession, labelTextPattern);

        public T[] FindByType<T>(bool showFound = false) => FindByType<T>(typeof(T).ToString(), showFound);

        public T[] FindByType<T>(string typeName, bool showFound = false) => Sap.FindByType<T>(GuiSession, typeName, showFound);

        public Grid FindGrid(string pathId)
        {
            return new Grid(this, pathId);
        }

        public Table FindTable(string pathId, int delayAfterScroll = 0)
        {
            return new Table(this, pathId);
        }

        public LabelTable FindLabelTable(string guiUsrControlPathId)
        {
            return new LabelTable(this, guiUsrControlPathId);
        }

        public void SelectComboBoxEntry(string guiComboBoxPathId, string entryToSelectRegex)
        {
            var guiComboBox = FindById<GuiComboBox>(guiComboBoxPathId);
            guiComboBox.Key = Rpa.COMCollectionToICollection<GuiComboBoxEntry>(guiComboBox.Entries).Where(x => Rpa.IsMatch(x.Value, entryToSelectRegex)).Select(x => x.Key).FirstOrDefault();
        }

        /// <summary>
        /// Create a new session using the existing session as creator.
        /// </summary>
        public void CreateNewSession() => CreateNewSession(this);
        /// <summary>
        /// Create a new session from another existing session. The created session goes to collection of opened sessions within App COM object.
        /// </summary>
        /// <param name="session">Session used to create the new one.</param>
        public void CreateNewSession(Session session)
        {
            Trace.WriteLine($"Creating a new session from Session[{session.Index}]...");
            session.GuiSession.CreateSession();
            Thread.Sleep(_sessionWaitMilisec); // wait otherwise new session cannot be captured
            //Log.MessageBox("New session created.");
            //Sap.MapExistingSessions();
            Trace.WriteLine(string.Join(Environment.NewLine,
                $"The connection after creating a new session and updating connection through interop engine:",
                Connection));
        }

        public static string SessionInfo(GuiSession session)
        {
            if (session == null) return "No Session";
            return string.Join(Environment.NewLine,
                $"    ID: \"{session.Id}\"",
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

        public bool ExistsById<T>(string pathId) => Sap.ExistsById<T>(GuiSession as GuiComponent, pathId);

        public bool ExistsById(string pathId) => Sap.ExistsById<dynamic>(GuiSession as GuiComponent, pathId);

        public bool ExistsByText<T>(string textRegex) => Sap.ExistsByText<T>(GuiSession as GuiComponent, textRegex);

        public bool ExistsByText(string textRegex) => Sap.ExistsByText<dynamic>(GuiSession as GuiComponent, textRegex);

        public bool ExistsTextInside(string parentPathId, string textRegex) => ExistsTextInside<dynamic,dynamic>(parentPathId, textRegex);

        /// <summary>
        /// Verifies if a text exists inside any descendant objects of a parent object which is descendant of this Session.GuiSession.
        /// </summary>
        /// <typeparam name="P">Type of the parent object.</typeparam>
        /// <typeparam name="C">Type of the object that contains a Text property in which the regex passed as parameter must match.
        /// Note that it isn't the type of the parent object, instead, it's the type of its child that must contain the regex like text.</typeparam>
        /// <param name="parentPathId">Path ID from the parent object. All its children are considered recursively (children of children, etc.)</param>
        /// <param name="textRegex">Pattern which any descendant's Text property must match</param>
        /// <returns>Boolean indicating if there is at least one descendant with the text specified.</returns>
        public bool ExistsTextInside<P, C>(string parentPathId, string textRegex) => Sap.ExistsTextInside<P, C>((GuiComponent)GuiSession, parentPathId, textRegex);

        public C[] FindTextInside<P, C>(string parentPathId, string textRegex) => Sap.FindTextInside<P, C>((GuiComponent)GuiSession, parentPathId, textRegex);

        public Grid NewGridView(string idGuiGridView) => NewGridView(FindById<GuiGridView>(idGuiGridView));
        public Grid NewGridView(GuiGridView guiGridView)
        {
            return new Grid(this, guiGridView);
        }

        public SapGuiObject[] AllDescendants() => Sap.AllDescendants(GuiSession);

        public string AllSessionIdsInfo() => Sap.AllSessionIdsInfo(GuiSession);

        public void ShowAllSessionIdsInfo() => Trace.WriteLine(AllSessionIdsInfo());

        public string CurrentStatusBarInfo()
        {
            return string.Join(Environment.NewLine,
                $"Current Status Bar properties:",
                $"  MessageType: {CurrentStatusBar.MessageType}",
                $"  Text: {CurrentStatusBar.Text}");
        }

        public void ShowCurrentStatusBarInfo()
        {
            Trace.WriteLine(CurrentStatusBarInfo());
        }
        public bool IsStatusType(StatusType status)
        {
            string statusLetter = StatusTypeEnum.GetStatusTypeLetter(status);
            if (Rpa.IsMatch(CurrentStatusBar.MessageType, statusLetter))
                return true;
            else
                return false;
        }

        public bool IsStatusMessage(string messageRegex)
        {
            if (Rpa.IsMatch(CurrentStatusBar.Text, messageRegex))
                return true;
            else
                return false;
        }

        /// <summary>
        /// Get session GuiFrameWindow by index.
        /// </summary>
        /// <param name="wndIndex">Index of the window. For example, to get the window of ID "/app/conn/wnd[0]" it should be 0.</param>
        /// <returns></returns>
        public GuiFrameWindow GetWindow(int wndIndex) => FindById<GuiFrameWindow>($"wnd[{wndIndex}]");

        /// <summary>
        /// Automate pressing enter for a number of times
        /// </summary>
        /// <param name="timesToPress">Number of times to press the enter key.</param>
        /// <param name="pressingIntervalMillisec">Interval in milliseconds between each key press.</param>
        /// <param name="wndIndex">GuiFrameWindow index of the window to apply the enter pressing action.</param>
        private void PressVKey(int vKeyNumber, int timesToPress = 1, int pressingIntervalMillisec = 0, int wndIndex = 0)
        {
            for (int i = 0; i < timesToPress; i++)
            {
                GetWindow(wndIndex).SendVKey(vKeyNumber);
                Thread.Sleep(pressingIntervalMillisec);
            }
        }

        public void PressEnter(int timesToPress = 1, int pressingIntervalMillisec = 0, int wndIndex = 0)
        {
            PressVKey(0, timesToPress, pressingIntervalMillisec, wndIndex);
        }

        public void PressEsc(int timesToPress = 1, int pressingIntervalMillisec = 0, int wndIndex = 0)
        {
            PressVKey(12, timesToPress, pressingIntervalMillisec, wndIndex);
        }

        /// <summary>
        /// Unconditionally close the SAP session GuiFrameWindow. Auto-processes any confirmation pop-up that might appear.
        /// </summary>
        /// <param name="confirmButtonTxt">Text in the button to confirm to close the Session window.</param>
        public void Close(string confirmButtonTxt = "Yes")
        {
            bool confirmationPopUpWillAppear = Connection.Sessions.Length == 1;

            CurrentFrameWindow.Close();

            if (confirmationPopUpWillAppear)
            {
                // Press "Yes" in confirmation pop-up.
                FindByText<GuiButton>(confirmButtonTxt).First().Press();
            }
        }
    }
}
