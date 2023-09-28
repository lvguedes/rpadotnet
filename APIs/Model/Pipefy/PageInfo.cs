using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy
{
    public class PageInfo
    {
        public string EndCursor { get; set; }
        public bool HasNextPage { get; set; }
        public bool HasPreviousPage { get; set; }
        public string StartCursor { get; set; }
    }
}
