using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class ExceededRetryLimitSapConnectionException : SapException
    {
        private static string defMsg = "SAP Connection failed and reached the retry limit.";
        public ExceededRetryLimitSapConnectionException() : base(defMsg) { }
        public ExceededRetryLimitSapConnectionException(string msg) : base(msg) { }
        //public ExceededRetryLimitSapConnectionException(string msg, Exception e) : base(msg, e) { }
    }
}
