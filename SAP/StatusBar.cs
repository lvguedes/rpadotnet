using RpaLib.ProcessAutomation;
using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class StatusBar
    {
        public Session Session { get; private set; }

        public SapComWrapper<GuiStatusbar> GuiStatusbar
        {
            get
            {
                return Session.FindById<GuiStatusbar>("wnd[0]/sbar");
            }
        }

        public string StatusLetter
        {
            get => GuiStatusbar.Com.MessageType;
        }
        public StatusType StatusType
        {
            get => GetStatusFromLetter(StatusLetter);
        }

        public string Text { get => Ut.Replace(GuiStatusbar.Text, @"\s+", " "); }

        public StatusBar(Session session)
        {
            Session = session;
        }

        public static StatusType GetStatusFromLetter(string statusLetter)
        {
            switch (statusLetter)
            {
                case "E": return StatusType.Error;
                case "W": return StatusType.Warning;
                case "S": return StatusType.Success;
                case "A": return StatusType.Abort;
                case "I": return StatusType.Information;
                default: return StatusType.None;
            }
        }

        public static string GetLetterFromStatus(StatusType statusType)
        {
            switch (statusType)
            {
                case StatusType.Error: return "E";
                case StatusType.Warning: return "W";
                case StatusType.Success: return "S";
                case StatusType.Abort: return "A";
                case StatusType.Information: return "I";
                default: return null;
            }
        }
    }
}
