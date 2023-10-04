using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public enum StatusType
    {
        Error,
        Warning,
        Success,
        Abort,
        Information
    }
    public class StatusTypeEnum
    {
        public static string GetStatusTypeLetter(StatusType statusType)
        {
            switch (statusType)
            {
                case StatusType.Error: return "E";
                case StatusType.Warning: return "W";
                case StatusType.Success: return "S";
                case StatusType.Abort: return "A";
                case StatusType.Information: return "I";
                default: return null;
            }
        }
    }
}
