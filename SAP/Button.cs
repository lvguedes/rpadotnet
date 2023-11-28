using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SapLegacy = RpaLib.SAP.Legacy.Sap;

namespace RpaLib.SAP
{
    public class Button : SapComponent<GuiButton>
    {
        public Button(Session session, string pathId) : base(session, pathId)
        { }

        public void Click()
        {
            Tracing.Log.Write($"Trying to click button {Name} ({Description}).");
            Click(PathId);
            Tracing.Log.Write($"Button \"{Name}\" ({Description}) clicked.");
        }
        public void Click(string fullPathId) => Session.FindById<GuiButton>(fullPathId).Press();
    }
}
