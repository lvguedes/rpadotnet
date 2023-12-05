using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.DataExtraction
{
    public class XPathResultsException : RpaLibException
    {
        public static string Msg { get; set; } = $"The XPath \"{Xml.XPathExpr}\" search found ";

        public XPathResultsException(string msg) : base(msg) { }
        public XPathResultsException(string msg, Exception e) : base(msg, e) { }
    }
    public class XPathReturnedZeroResultsException : XPathResultsException
    {
        public XPathReturnedZeroResultsException() : base(Msg + "nothing") { }
        public XPathReturnedZeroResultsException(string msg) : base(msg) { }
        public XPathReturnedZeroResultsException(string msg, Exception e) : base(msg, e) { }
    }

    public class XPathReturnedMoreThanOneResultException : XPathResultsException
    {
        public XPathReturnedMoreThanOneResultException() : base(Msg + "more than one result") { }
        public XPathReturnedMoreThanOneResultException(string msg) : base(msg) { }
        public XPathReturnedMoreThanOneResultException(string msg, Exception e) : base(msg, e) { }
    }

    public class XPathReturnedEmptyStringResultException : XPathResultsException
    {
        public XPathReturnedEmptyStringResultException() : base(Msg + "empty string as result") { }
        public XPathReturnedEmptyStringResultException(string msg) : base(msg) { }
        public XPathReturnedEmptyStringResultException(string msg, Exception e) : base(msg, e) { }
    }
}