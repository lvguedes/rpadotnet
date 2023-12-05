using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class PatternNotFoundForSelectionException : SapException
    {
        public PatternNotFoundForSelectionException(string pattern) : base(GetMsg(pattern)){ }

        private static string GetMsg(string pattern)
        {
            return $"Item not found for selection. Regex pattern used: /{pattern}/";
        }
    }
}
