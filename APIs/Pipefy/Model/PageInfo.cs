using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class PageInfo
    {
        public string Endcursor { get; set; }
        public bool Hasnextpage { get; set; }
        public bool Haspreviouspage { get; set; }
        public string Startcursor { get; set; }
    }
}
