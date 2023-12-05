using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Exceptions
{
    public class RpaLibTimeoutException : RpaLibException
    {
        const string _defaultMessage = "The operation has timed out.";
        public RpaLibTimeoutException() : base(_defaultMessage) { }
        public RpaLibTimeoutException(string message) : base(message) { }
        public RpaLibTimeoutException(string message, Exception innerException) : base(message, innerException) { }
    }
}
