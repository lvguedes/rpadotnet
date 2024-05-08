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
        private string jwt = "Bearer eyJ0eXAiOiJK...";
        private string pipeId = "902791";

        private Pipefy GetPipefyWithouPipeId()
        {
            return new Pipefy(jwt);
        }

        private Pipefy GetPipefyWithPipeId()
        {
            var pipeId = "902791";
            Pipefy pipefy = new Pipefy(jwt, pipeId);

            return pipefy;
        }

        [Fact]
        public void QueryAccountInfo()
        {
            var pipefy = GetPipefyWithouPipeId();

            string logMsg = $"Me:\n{pipefy.QueryUserInfo()}\nOrganizations:\n{pipefy.QueryOrganizations()}";

            Console.WriteLine(logMsg);
        }

        [Fact]
        public void QueryPipeFields()
        {
            var pipefy = GetPipefyWithPipeId();

            var pipeFields = pipefy.QueryStartFormFields();
        }

        [Fact]
        public void ShowInfo()
        {
            var pipefy = GetPipefyWithPipeId();
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
            var pipefy = GetPipefyWithouPipeId();

            pipefy.ShowInfo(PipefyInfo.Organizations);
        }

        [Fact]
        public async void NonBlockingShowOrgInfo()
        {
            var pipefy = GetPipefyWithouPipeId();

            var result = await pipefy.ShowInfoAsync(PipefyInfo.Organizations);

            Console.WriteLine(result);
        }

        [Fact]
        public void UploadFile()
        {
            var pipefy = GetPipefyWithPipeId();

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

        [Fact]
        public async Task ExportPipeReportAsync()
        {
            var pipefy = GetPipefyWithPipeId();
            var reportNameRegex = @"Automação Cronograma de Fechamento - Visualização Sharepoint";

            var downloadUrl = await pipefy.ExportPipeReportAsync(reportNameRegex, @"%USERPROFILE%\Downloads");
        } 

    }
}
