using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class PhaseField : IPipefyField
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public object Value { get; set; }
    }
}
