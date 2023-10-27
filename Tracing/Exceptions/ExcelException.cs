using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Tracing.Exceptions
{
    public abstract class ExcelException : Exception
    {
        public ExcelException(string message) : base(message) { }
    }
}
