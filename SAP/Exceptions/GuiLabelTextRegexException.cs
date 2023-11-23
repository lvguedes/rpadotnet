using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class GuiLabelTextRegexException : SapException
    {
        public GuiLabelTextRegexException(string regexPattern)
            : base(GetMsg(regexPattern)) { }

        private static string GetMsg(string regexPattern)
        {
            return $"The regex \"{regexPattern}\" didn't match any GuiLabel.Text";
        }
    }
}
