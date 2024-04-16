using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class SapTextFieldException : SapException
    {
        public SapTextFieldException(string inputText) 
            : base(GetMsg(inputText)) { }

        public SapTextFieldException(string inputText, Exception innerException)
            : base(GetMsg(inputText), innerException) { }

        private static string GetMsg(string inputText) => 
            $"It wasn't possible to write '{inputText}' to the SAP text field.";

    }
}
