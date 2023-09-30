using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    using CardEdges = List<CardEdge>;
    public class CardConnection
    {
        public CardEdges Edges { get; set; }
        public PageInfo Pageinfo { get; set; }
    }
}
