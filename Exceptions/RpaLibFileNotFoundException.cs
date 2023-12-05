using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Exceptions
{
    public class RpaLibFileNotFoundException : RpaLibException
    {
        const string _defaultMessage = "Unable to find the specified file.";
        public RpaLibFileNotFoundException() : base(_defaultMessage) { }
        public RpaLibFileNotFoundException(string message) : base(message) { }
        public RpaLibFileNotFoundException(string message, Exception innerException) : base(message, innerException) { }
    }
}
