using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Model
{
    public class SapGuiObject
    {
        public string PathId { get; set; }
        public string Type { get; set; }
        public dynamic Obj { get; set; }
        public string Text { get; set; }
    }
}
