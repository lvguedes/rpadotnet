using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class ExceededRetryLimitSapConnectionException : Exception
    {
        private static string defMsg = "SAP Connection failed and reached the retry limit.";
        public ExceededRetryLimitSapConnectionException() : base(defMsg) { }
        public ExceededRetryLimitSapConnectionException(string msg) : base(msg) { }
        public ExceededRetryLimitSapConnectionException(string msg, Exception e) : base(msg, e) { }
    }

    public class FieldDatatypeConversionException : Exception
    {
        private static string defMsg = @"Could not evaluate the correct type when trying to convert from string";
        public FieldDatatypeConversionException() : base(defMsg) { }
        public FieldDatatypeConversionException(string msg) : base(msg) { }
        public FieldDatatypeConversionException(string msg, Exception e) : base(msg, e) { }
    }
}
