using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Exceptions
{
    public class RpaLibArgumentNullException : RpaLibException
    {
        public RpaLibArgumentNullException(string message) : base(message) { }
        public RpaLibArgumentNullException(string message, Exception innerException) : base(message, innerException) { }
    }
}
