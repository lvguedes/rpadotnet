using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class CardField : IPipefyField
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public object Value { get; set; }
    }
}
