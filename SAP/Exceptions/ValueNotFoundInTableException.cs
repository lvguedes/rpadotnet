using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class ValueNotFoundInTableException : SapException
    {
        public ValueNotFoundInTableException(string message) : base(message) { }
    }
}
