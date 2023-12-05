using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Exceptions
{
    internal class RpaLibArgumentException : RpaLibException
    {
        public RpaLibArgumentException(string message) : base(message) { }
        public RpaLibArgumentException(string message, Exception innerException) : base(message, innerException) { }
    }
}
