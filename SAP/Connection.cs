using Microsoft.Office.Interop.Excel;
using RpaLib.SAP.Legacy;
using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class Connection
    {
        public GuiConnection GuiConnection { get; private set; }


        public string Id
        {
            get => GuiConnection.Id;
        }

        public string ConnectionString
        {
            get => GuiConnection.ConnectionString;
        }

        public string Description
        {
            get => GuiConnection.Description;
        }

        public Session[] Sessions
        {
            get => GetSessions();
        }

        public bool DisabledByServer
        {
            get => GuiConnection.DisabledByServer;
        }

        public Connection(GuiConnection guiConnection)
        {
            GuiConnection = guiConnection;
        }

        public Connection Update()
        {
            var connectionFound = App.FindConnectionById(Id);
            GuiConnection = connectionFound.GuiConnection;

            return connectionFound;
        }

        public Session[] GetSessions()
        {
            List<Session> sessions = new List<Session>();
            foreach (GuiSession guiSession in  GuiConnection.Sessions)
            {
                sessions.Add(new Session(guiSession, this));
            }

            return sessions.ToArray();
        }

        public string ConnectionInfo() => ConnectionInfo(this);
        public static string ConnectionInfo(Connection connection)
        {
            return string.Join(Environment.NewLine, connection);
        }

        public override string ToString()
        {
            return string.Join(Environment.NewLine,
                $"  Connection:",
                $"    Id: \"{Id}\"",
                $"    Description: \"{Description}\"",
                $"    ConnectionString: \"{ConnectionString}\"",
                $"    Sessions: \"{SessionsListTransaction()}\"",
                $"    DisabledByServer: \"{DisabledByServer}\"",
                $"",
                $"  The Sessions/Children elements:",
                SessionsInfo()
                );
        }

        public string SessionsListTransaction()
        {
            StringBuilder sb = new StringBuilder();
            foreach(var session in Sessions)
            {
                if (!string.IsNullOrEmpty(sb.ToString()))
                    sb.Append(", ");
                sb.Append($"{{ Id: \"{Id}\"; Transaction:\"{session.Transaction}\" }}");
            }
            return sb.ToString();
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
    }
}
