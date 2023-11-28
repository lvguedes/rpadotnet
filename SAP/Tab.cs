using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SapLegacy = RpaLib.SAP.Legacy.Sap;

namespace RpaLib.SAP
{
    public class Tab : SapComponent<GuiTab>
    {
        public Session Session { get; private set; }
        public int Index { get; set; }
        public Field[] Fields { get; set; }

        public Tab(Session session, string pathId) : base(session, pathId)
        {
            Session = session;
        }

        public void Select() => Select(PathId);
        public void Select(string fullPathId)
        {
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Trying to select tab \"{Name}\"",
                $"({fullPathId})"));
            Session.FindById<GuiTab>(fullPathId).Com.Select();
            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Tab \"{Name}\" selected.",
                $"{fullPathId}"));
        }
    }
}
