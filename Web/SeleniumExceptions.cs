using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.Web
{
    public class TabNotFoundException : RpaLibException
    {
        private static string defMsg = "The tab wasn't found.";
        private static string defMsgWithTab = "The tab \"{0}\" wasn't found.";
        public TabNotFoundException() : base(defMsg) { }
        public TabNotFoundException(string tabName) : base(string.Format(defMsgWithTab, tabName)) { }
        public TabNotFoundException(string tabName, Exception e) : base(string.Format(defMsgWithTab, tabName)) { }
    }
    public class ChromeDriverDeprecatedException : RpaLibException
    {
        private static string defMsg = "Chrome driver version is deprecated and doesn't work with the Chrome system's current version.";

        public ChromeDriverDeprecatedException(string msg) : base(string.Concat(defMsg, " ", msg)) { }
        public ChromeDriverDeprecatedException(string msg, Exception ex) : base(msg, ex) { }
    }
}
