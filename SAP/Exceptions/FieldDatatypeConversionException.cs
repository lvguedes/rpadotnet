using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class FieldDatatypeConversionException : SapException
    {
        private static string defMsg = @"Could not evaluate the correct type when trying to convert from string";
        public FieldDatatypeConversionException() : base(defMsg) { }
        public FieldDatatypeConversionException(string msg) : base(msg) { }
        //public FieldDatatypeConversionException(string msg, Exception e) : base(msg, e) { }
    }
}
