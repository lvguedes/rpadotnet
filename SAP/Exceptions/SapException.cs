using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.SAP.Exceptions
{
    public abstract class SapException : RpaLibException
    {
        public SapException(string message) : base(message) { }
        public SapException(string message, Exception innerException) : base(message, innerException) { }
    }
}
