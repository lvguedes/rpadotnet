using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using SapLegacy = RpaLib.SAP.Legacy.Sap;

namespace RpaLib.SAP
{
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

        public void AccessTransaction(string transactionId) => SapLegacy.AccessTransaction(this, transactionId);
        public T FindById<T>(string id) => (T)FindById(id);

        public T[] FindByText<T>(string labelText) => SapLegacy.FindByText<T>(GuiSession, labelText);

        public void CreateNewSession() => CreateNewSession(this);
        public void CreateNewSession(Session session)
        {
            Tracing.Log.Write($"Creating a new session from Session[{session.Index}]...");
            session.GuiSession.CreateSession();
            Thread.Sleep(2000); // wait otherwise new session cannot be captured
            //Log.MessageBox("New session created.");
            SapLegacy.UpdateConnections();
            SapLegacy.MapExistingSessions();
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"The connection after creating a new session and updating connection through interop engine:",
                SapLegacy.ConnectionInfo()));
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
}
