using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
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
}
