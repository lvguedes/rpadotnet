using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP.Exceptions
{
    public class ConnectionNotUniqueException : SapException
    {
        public ConnectionNotUniqueException(int numberOfConnections) : base(GetMsg(numberOfConnections)) { }

        private static string GetMsg(int nConn)
        {
            return $"There is not only one connection. Number of connections found within COM: \"{nConn}\"";
        }
    }
}
