using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class OptionNotFoundException : SapException
    {
        public OptionNotFoundException(string message, Exception ex) : base(message, ex) { }
    }
}
