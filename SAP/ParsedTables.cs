using RpaLib.SAP.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class ParsedTables
    {
        public Label[][] GuiLabels { get; private set; }
        public DataTable DataTable { get; private set; }
        public bool Selected { get; private set; }

        public ParsedTables(Label[][] guiLabels = null, DataTable dataTable = null, bool selected = false)
        {
            GuiLabels = guiLabels;
            DataTable = dataTable;
            Selected = selected;
        }
    }
}
