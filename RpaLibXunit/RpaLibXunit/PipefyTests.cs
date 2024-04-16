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

    }
}
