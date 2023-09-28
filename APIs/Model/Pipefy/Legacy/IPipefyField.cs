using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public interface IPipefyField : IPipefyObject
    {
        string Name { get; set; }
        object Value { get; set; }
    }
}
