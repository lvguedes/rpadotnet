using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using RpaLib.APIs.Pipefy;

namespace RpaLibXunit
{
    public class PipefyTests
    {
        private string jwt = "Bearer {enter token here}";

        [Fact]
        public void QueryAccountInfo()
        {
            Pipefy pipefy = new Pipefy(jwt);

            string logMsg = $"Me:\n{pipefy.QueryUserInfo()}\nOrganizations:\n{pipefy.QueryOrganizations()}";

            Console.WriteLine(logMsg);
        }

        [Fact]
        public void BlockingShowOrgInfo()
        {
            Pipefy pipefy = new Pipefy(jwt);

            pipefy.ShowInfo(PipefyInfo.Organizations);
        }

        [Fact]
        public async void NonBlockingShowOrgInfo()
        {
            Pipefy pipefy = new Pipefy(jwt);

            var result = await pipefy.ShowInfoAsync(PipefyInfo.Organizations);

            Console.WriteLine(result);
        }

    }
}
