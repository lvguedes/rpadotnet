using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.Tracing.Exceptions
{
    public abstract class ExcelException : RpaLibException
    {
        public ExcelException(string message) : base(message) { }
    }
}
