using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using RpaLib.APIs.Pipefy;
using RpaLib.ProcessAutomation;
using System.Diagnostics;

namespace RpaLibXunit
{
    public class PipefyTests
    {
        //private string jwt = "Bearer {enter token here}";
        private string jwt = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJ1c2VyIjp7ImlkIjo5MDU5NDYsImVtYWlsIjoiYW5kcmVhLmpld29yb3dza2lAY2FwZ2VtaW5pLmNvbSIsImFwcGxpY2F0aW9uIjozMDAxNzQ2MDJ9fQ.xT705GbX3CbTv8gQ0WH6kOPcsr-2DNa1tTTsPOXLdeQFyF7N55t2r-F-MZzCGZvMQi7pBZOiFbBsu-83rNm6lQ";

        [Fact]
        public void QueryAccountInfo()
        {
            Pipefy pipefy = new Pipefy(jwt);

            string logMsg = $"Me:\n{pipefy.QueryUserInfo()}\nOrganizations:\n{pipefy.QueryOrganizations()}";

            Console.WriteLine(logMsg);
        }

        [Fact]
        public void QueryPipeFields()
        {
            var pipeId = "902791";
            Pipefy pipefy = new Pipefy(jwt, pipeId);

            var pipeFields = pipefy.QueryStartFormFields();
        }

        [Fact]
        public void ShowInfo()
        {
            var pipeId = "902791";
            Pipefy pipefy = new Pipefy(jwt, pipeId);
            var phases = new string[] { "6118794", "6118797", "7287878", "6118796", "6118799", "7287890", "7287891", "7335279" };

            pipefy.ShowInfo(PipefyInfo.PhasesAndCardsCount);

            foreach (var phaseId in phases)
            {
                Trace.WriteLine($"\nPhase fields from phase: {phaseId}");
                pipefy.ShowInfo(PipefyInfo.PhaseFields, phaseId);
            }
            
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

        [Fact]
        public async void UploadFile()
        {
            var pipeId = "902791";
            Pipefy pipefy = new Pipefy(jwt, pipeId);

            var presignedUrl = pipefy.CreatePresignedUrl("ArquivoTeste.xlsx").Data.CreatePresignedUrl;
            var upload = presignedUrl.Url;
            var download = presignedUrl.DownloadUrl;

            var testFilePath = Ut.GetFullPath(@"%USERPROFILE%\Downloads\Test.xlsx");

            //Task.Run(async () =>
            //{
            //    await Ut.HttpPutFileAsync(testFilePath, upload);
            //});

            var t1 = Ut.HttpPutFileAsync(testFilePath, upload);

            Task.WaitAll(t1);

            Console.WriteLine(t1.Result.StatusCode);
            Console.WriteLine($"The download URL is: {download}");
        }

    }
}
