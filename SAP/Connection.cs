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
        public GuiConnection GuiConnection
        {
            get => App.FindGuiConnectionByDesc(Description);
        }


        public string Id { get; private set; }

        public string ConnectionString { get; private set; }

        public string Description { get; private set; }

        public Session[] Sessions
        {
            get => GetSessions();
        }

        public Connection(GuiConnection guiConnection)
        {
            Id = guiConnection.Id;
            ConnectionString = guiConnection.ConnectionString;
            Description = guiConnection.Description;
        }

        //public Connection Update()
        //{
        //    var connectionFound = App.FindConnectionById(Id);
        //    //GuiConnection = connectionFound.GuiConnection;

        //    return connectionFound;
        //}

        public Session[] GetSessions()
        {
            List<Session> sessions = new List<Session>();
            foreach (GuiSession guiSession in  GuiConnection.Sessions)
            {
                sessions.Add(new Session(guiSession, GuiConnection));
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
                $"    Id: \"{GuiConnection.Id}\"",
                $"    Description: \"{GuiConnection.Description}\"",
                $"    ConnectionString: \"{GuiConnection.ConnectionString}\"",
                $"    Sessions: \"{SessionsListTransaction()}\"",
                $"    DisabledByServer: \"{GuiConnection.DisabledByServer}\"",
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

        public void Close()
        {
            GuiConnection.CloseConnection();
        }
    }
}
