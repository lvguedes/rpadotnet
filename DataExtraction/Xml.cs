using RpaLib.Tracing;
using RpaLib.ProcessAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RpaLib.DataExtraction
{
    public class Xml
    {
        public static string XPathExpr { get; private set; }
        public string Path { get; set; }
        public XElement XElement { get; private set; }

        public Xml() { }
        public Xml(string path)
        {
            Path = Ut.GetFullPath(path);
        }

        public string XPath(string xpath, bool relaxed = false)
        {
            XPathExpr = xpath;
            string searchResult = string.Empty;

            Log.Write($"Trying to find xpath \"{xpath}\" inside file \"{Path}\"");
            try
            {
                searchResult = XPath(Path, xpath);
            }
            catch (Exception ex) when (
                ex is XPathReturnedEmptyStringResultException ||
                ex is XPathReturnedMoreThanOneResultException ||
                ex is XPathReturnedZeroResultsException
                )
            {
                if (!relaxed)
                    throw ex;
            }
 
            Log.Write($"Found value (\"{xpath}\" in file \"{Path}\"): {searchResult}");

            return searchResult;
        }

        public static IEnumerable<XElement> XPathElt(XNode node, string xpath) => System.Xml.XPath.Extensions.XPathSelectElements(node, xpath);

        public static IEnumerable<XElement> XPathElt(string xmlPath, string xpath)
        {
            XElement xml = XElement.Load(xmlPath);
            return XPathElt(xml, xpath);
        }

        public static string XPath(string xmlPath, string xpath)
        {
            IEnumerable<XElement> found = XPathElt(xmlPath, xpath);

            if (found.Count() == 0)
            {
                throw new XPathReturnedZeroResultsException();
            }
            else if (found.Count() > 1)
            {
                throw new XPathReturnedMoreThanOneResultException();
            }
            else if (string.IsNullOrEmpty(found.ElementAt<XElement>(0).Value))
            {
                throw new XPathReturnedEmptyStringResultException();
            }
            else
            {
                return found.ElementAt<XElement>(0).Value;
            }
        }
    }
}
