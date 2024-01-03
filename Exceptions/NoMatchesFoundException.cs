using RpaLib.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.ProcessAutomation.Exceptions
{
    public class NoMatchesFoundException : RpaLibException
    {
        public NoMatchesFoundException(string message) : base(message) { }
    }
}
