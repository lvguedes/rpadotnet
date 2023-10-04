using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SDKs.AutoIt.Exceptions
{
    public class ProcessNotFoundException : Exception
    {
        public ProcessNotFoundException(string psNameOrPid) : base(GetMessage(psNameOrPid)) { }
        public ProcessNotFoundException(string psNameOrPid, Exception innerException) : base(GetMessage(psNameOrPid), innerException) { }

        private static string GetMessage(string psNameOrPid)
        {
            return $"Process with name or PID \"{psNameOrPid}\" was not found.";
        }
    }
}
