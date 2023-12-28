using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class PopupNotFoundException : SapException
    {
        public PopupNotFoundException(string message) : base(message) { }
    }
}
