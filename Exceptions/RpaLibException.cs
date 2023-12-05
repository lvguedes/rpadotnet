using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Exceptions
{
    public class RpaLibException : Exception
    {
        public RpaLibException(string message) : base(message) { }
        public RpaLibException(string message, Exception innerException) : base(message, innerException) { }
    }
}
