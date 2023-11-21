using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public abstract class SapException : Exception
    {
        public SapException(string message) : base(message) { }
    }
}
