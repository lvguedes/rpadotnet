using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Model
{
    public class Label
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public string Text { get; set; }
        public SapComWrapper<GuiLabel> GuiLabel { get; set; }
    }
}
