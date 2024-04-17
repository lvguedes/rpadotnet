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
        [Fact]
        public void QueryAccountInfo()
        {
            string jwt = "Bearer eyJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJQaXBlZnkiLCJpYXQiOjE2OTUyMjI1MTksImp0aSI6ImM4NzRhN2EwLTVlM2QtNDUyZS1iOGQ3LTBlZGM3NTI1ZDFlZCIsInN1YiI6MzAxMjc3MzAyLCJ1c2VyIjp7ImlkIjozMDEyNzczMDIsImVtYWlsIjoiam9zZS5yYWltdW5kb0BjYXBnZW1pbmkuY29tIiwiYXBwbGljYXRpb24iOjMwMDI3Nzc2NSwic2NvcGVzIjpbXX0sImludGVyZmFjZV91dWlkIjpudWxsfQ.LpUVnEY-zwODf6jB_MLUao8-HMC0CBMvJsnT2Na2ILjDjAKqYBQ6fv-00YzmSwh5Z9gg6t_dJEABj7srFTag7g";
            Pipefy pipefy = new Pipefy(jwt);

            string logMsg = $"Me:\n{pipefy.QueryUserInfo()}\nOrganizations:\n{pipefy.QueryOrganizations()}";

            Console.WriteLine(logMsg);
        }

        [Fact]
        public void BlockingShowOrgInfo()
        {
            string jwt = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9" +
                ".eyJ1c2VyIjp7ImlkIjo5MDU5NDYsImVtYWlsIjoiYW5kcmVhLmpld29yb3dza2lAY2FwZ2VtaW5pLmNvbSIsImFwcGxpY2F0aW9uIjozMDAxNzQ2MDJ9fQ" +
                ".xT705GbX3CbTv8gQ0WH6kOPcsr-2DNa1tTTsPOXLdeQFyF7N55t2r-F-MZzCGZvMQi7pBZOiFbBsu-83rNm6lQ";

            Pipefy pipefy = new Pipefy(jwt);

            pipefy.ShowInfo(PipefyInfo.Organizations);
        }

        [Fact]
        public async void NonBlockingShowOrgInfo()
        {
            string jwt = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9" +
                ".eyJ1c2VyIjp7ImlkIjo5MDU5NDYsImVtYWlsIjoiYW5kcmVhLmpld29yb3dza2lAY2FwZ2VtaW5pLmNvbSIsImFwcGxpY2F0aW9uIjozMDAxNzQ2MDJ9fQ" +
                ".xT705GbX3CbTv8gQ0WH6kOPcsr-2DNa1tTTsPOXLdeQFyF7N55t2r-F-MZzCGZvMQi7pBZOiFbBsu-83rNm6lQ";

            Pipefy pipefy = new Pipefy(jwt);

            var result = await pipefy.ShowInfoAsync(PipefyInfo.Organizations);

            Console.WriteLine(result);
        }

    }
}
