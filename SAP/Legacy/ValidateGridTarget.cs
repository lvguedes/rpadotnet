using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Legacy
{
    public static class ValidateGridTarget
    {
        public static void ValidateTarget(string id, GuiGridView gridView, GuiSession session)
        {
            if (gridView == null && string.IsNullOrEmpty(id))
                throw new Exception("Parameters for the Target object not provided: id or GuiGridView object");

            if (session == null)
                throw new Exception("SAP session parameter is required and was not provided.");
        }
    }
}
