using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class SapInSessionException : SapException
    {
        public StatusBar CurrentStatusBar { get; private set; }

        public SapInSessionException(StatusBar statusBar, Exception innerException) : base(innerException.Message, innerException) 
        {
            CurrentStatusBar = statusBar;
        }
    }
}
