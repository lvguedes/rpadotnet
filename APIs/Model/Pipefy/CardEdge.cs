using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy
{
    public class CardEdge
    {
        public string Cursor { get; set; }
        public Card Node { get; set; }
    }
}
