using Xunit;
using RpaLib.ProcessAutomation;
using System.Threading.Tasks;
using System;
using System.IO;

namespace RpaLibXunit
{
    public class Pipelines
    {
        [Fact]
        public void InterProcessCommunication()
        {
            var serverProjName = "IPCPipelineServer";
            var clientProjName = "IPCPipelineClient";

            Func<string, string, string, string> getPathWithName = (a, b, e) => Path.Combine("..", "..", "..", "..", "..", a, "bin", "Debug", $"{b}.{e}");
            Func<string, string> getExePath = (n) => getPathWithName(n, n, "exe");
            Func<string, string> getConfigPath = (n) => getPathWithName(n, "Config", "yaml");

            var serverOut = Ut.RunPromptCommandAsync(getExePath(serverProjName), $"\"{getConfigPath(serverProjName)}\"");
            var clientOut = Ut.RunPromptCommandAsync(getExePath(clientProjName), $"\"{getConfigPath(clientProjName)}\"");

            Task.WaitAll(clientOut);

            Assert.Contains("SIGTERM", serverOut?.Result.StdOut);
        }
    }
}