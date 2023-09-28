using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
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
}
