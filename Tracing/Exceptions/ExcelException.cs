using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.Tracing.Exceptions
{
    public class ExcelException : RpaLibException
    {
        public ExcelException(string message) : base(message) { }
        public ExcelException(string messsage, Exception innerException) : base(messsage, innerException) { }
    }
}
