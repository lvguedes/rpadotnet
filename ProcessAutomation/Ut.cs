using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Threading;
using System.Text.RegularExpressions;
//using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Data;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization;
using System.Net.Http;
using Newtonsoft.Json;
using RpaLib.Exceptions;
using FormsMsgBox = System.Windows.Forms.MessageBox;
using DialogResult = System.Windows.Forms.DialogResult;
using MessageBoxButtons = System.Windows.Forms.MessageBoxButtons;
using System.Net.Http.Headers;
using Process = System.Diagnostics.Process;
using ProcessStartInfo = System.Diagnostics.ProcessStartInfo;
using Stopwatch = System.Diagnostics.Stopwatch;
using RpaLib.Tracing;

/*
Reference Assemblies:
  - System.IO.Compression.FileSystem
*/


/// Generic RPA functions
namespace RpaLib.ProcessAutomation
{
    public delegate void VoidFuncVoid();
    public static class Ut
    {
        public static readonly Dictionary<string, string> CYGPATH =
            new Dictionary<string, string>() { { "PATH", GetFullPath(@"ProcessAutomation\lib\cygbase") } };

        const string CURLPATH = @"ProcessAutomation\bin";

        [DllImport("Kernel32.dll", SetLastError = true)]
        public static extern int SetStdHandle(int device, IntPtr handle);

        #region RegularExpressions

        /// <summary>
        /// Replace a regex pattern match with a replacement string.
        /// </summary>
        /// <param name="input">The input string.</param>
        /// <param name="pattern">The Regular Expression pattern.</param>
        /// <param name="replacement">The string to replace the pattern found.</param>
        /// <param name="regexOptions">RegexOptions enum to set regex options.</param>
        /// <returns>The input string with replacement applied.</returns>
        public static string Replace(string input, string pattern, string replacement,
            RegexOptions regexOptions = RegexOptions.IgnoreCase) => Regex.Replace(input, pattern, replacement);
        
        /// <summary>
        /// Search for the first occurrence of pattern within input string.
        /// </summary>
        /// <param name="input">Text where to search for.</param>
        /// <param name="pattern">Regex pattern representing what to look for.</param>
        /// <param name="regexOptions">DotNet RegexOptions controling the Regular Expression options.</param>
        /// <returns>The text found.</returns>
        public static string Match(string input, string pattern, RegexOptions regexOptions = RegexOptions.IgnoreCase) => Regex.Match(input, pattern, regexOptions).Value;

        /// <summary>
        /// Check if the regex pattern matches somewhere in input string.
        /// </summary>
        /// <param name="input">Input text where to search for.</param>
        /// <param name="pattern">Regex pattern to look for within input text.</param>
        /// <param name="regexOptions">RegexOptions from dotnet controling regex options.</param>
        /// <returns>True if pattern was found within input text, false otherwise.</returns>
        public static bool IsMatch(string input, string pattern, RegexOptions regexOptions = RegexOptions.IgnoreCase) => Regex.IsMatch(input, pattern, regexOptions);

        #endregion

        #region FileSystem
        /// <summary>
        /// Waits until a file (full path string) appears within its base directory.
        /// It will wait until file appear or timeout exceed. The refresh rate sleep 
        /// is also configurable (defaults to 0.5 seconds).
        /// </summary>
        /// <param name="fileFullPath">The full path to the file it's needed to wait until appear</param>
        /// <param name="timeoutSeconds">The timeout in seconds you need to wait</param>
        /// <param name="refreshDelay">The time in milliseconds between each refresh in files parent dir </param>
        public static void WaitUntilFileAppear(string fileFullPath, int timeoutSeconds, int refreshDelay = 500)
        {
            string parentDirFullPath = Path.GetDirectoryName(fileFullPath);
            string fileBasename = Path.GetFileName(fileFullPath);

            WaitUntilFileAppear(parentDirFullPath, fileBasename, timeoutSeconds, refreshDelay);
        }
        /// <summary>
        /// Waits until a file that matches the Regex appear in the full path to its parent dir given
        /// </summary>
        /// <param name="parentDirFullPath">File's parent directory full path</param>
        /// <param name="fileRegex">Regex to match the file name</param>
        /// <param name="timeoutSeconds">Timeout representing the maximuma time to wait for file until FileNotFoundException be thrown</param>
        /// <param name="refreshDelay">The time in milliseconds between each refresh in file's parent directory</param>
        /// <param name="regexOpts">Regular Expression matching options to be considered when matching, such as Ignoring Case</param>
        /// <exception cref="FileNotFoundException"></exception>
        public static void WaitUntilFileAppear(string parentDirFullPath, string fileRegex, int timeoutSeconds, int refreshDelay = 500,
            RegexOptions regexOpts = RegexOptions.IgnoreCase)
        {
            double waited;

            for (waited = 0; waited <= timeoutSeconds; waited += (double)refreshDelay / 1000)
            {
                if (IsFileInDir(fileRegex, parentDirFullPath))
                    return;
                Thread.Sleep(refreshDelay);
            }

            throw new RpaLibFileNotFoundException(parentDirFullPath + Path.DirectorySeparatorChar + fileRegex);
        }

        public static bool IsFileInDir(string fileNameRegex, string dirPath)
        {
            string dirFullPath = GetFullPath(dirPath);

            foreach (string file in Directory.GetFiles(dirFullPath))
            {
                if (Ut.IsMatch(file, fileNameRegex))
                {
                    return true;
                }
            }

            return false;
        }

        public static string[] MatchAllFiles(string fileNameRegex, string dirPath)
        {
            string dirFullPath = GetFullPath(dirPath);

            return Directory.GetFiles(dirFullPath).Cast<string>().Where(x => IsMatch(x, fileNameRegex)).ToArray();
        }

        public static string GetFullPath(string partialOrWithEnvVarsPath) =>
            Path.GetFullPath(Environment.ExpandEnvironmentVariables(partialOrWithEnvVarsPath));

        public static void Unzip(string sourceArchiveName, string destinationDirectoryName)
        {
            string sourceFile = GetFullPath(sourceArchiveName);
            string destDir = GetFullPath(destinationDirectoryName);
            ZipFile.ExtractToDirectory(sourceFile, destDir);
        }

        public static void MoveFileForced(string sourcePath, string destPath)
        {
            string sourceFullPath = GetFullPath(sourcePath);
            string destFullPath = GetFullPath(destPath);

            /*
            foreach (string file in MatchAllFiles(Path.GetFileName(sourceFullPath), Path.GetDirectoryName(destFullPath)))
                File.Delete(file);
            */

            if (IsFileInDir(Path.GetFileName(sourceFullPath), Path.GetDirectoryName(destFullPath)))
                File.Delete(Path.Combine(Path.GetDirectoryName(destFullPath), Path.GetFileName(sourceFullPath)));

            File.Move(sourceFullPath, destFullPath);
        }

        public static void CreateFileIfNotExists(string filePath)
        {
            var fullFilePath = Ut.GetFullPath(filePath);

            if (!File.Exists(fullFilePath))
                File.Create(fullFilePath);
        }

        #endregion

        #region Processes

        public static void WaitProcessStart(string processName, int timeoutSeconds = 300, int refreshDelayMillisec = 0, bool outputProcesses = false)
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            while (true)
            {
                foreach (Process proc in Process.GetProcesses())
                {
                    if (outputProcesses == true) Trace.WriteLine($"Process {proc.ProcessName}");
                    if (proc.ProcessName == processName)
                    {
                        return;
                    }
                    else if (timer.ElapsedMilliseconds == timeoutSeconds * 1000)
                    {
                        throw new RpaLibTimeoutException($"Timeout expired waiting for process \"{processName}\" for \"{timeoutSeconds}\" seconds.");
                    }
                    else
                    {
                        Thread.Sleep(refreshDelayMillisec);
                    }
                }
            }
        }

        public static void ApplyIfProcessExists(dynamic f, string processName, params dynamic[] parameters)
        {
            while (true)
            {
                foreach (Process proc in Process.GetProcesses())
                {
                    if (proc.ProcessName == processName)
                    {
                        f.Invoke(parameters);
                    }
                }
            }
        }

        public static bool ProcessExists(string processNamePattern)
        {
            foreach (var proc in Process.GetProcesses())
            {
                if (IsMatch(proc.ProcessName, processNamePattern))
                    return true;
            }
            return false;
        }

        public static void KillProcess(string processName, bool outputProcesses = false)
        {
            foreach (Process proc in Process.GetProcessesByName(processName))
            {
                if (outputProcesses) Trace.WriteLine($"Killing process {proc.ProcessName}");
                proc.Kill();
            }
        }

        public static void StartWaitProcess(string exePath, int timeoutSeconds = 300, bool outputProcesses = false)
        {
            string fullPath = Path.GetFullPath(exePath);
            string name = Path.GetFileNameWithoutExtension(fullPath);

            StartWaitProcess(fullPath, name, timeoutSeconds, outputProcesses);
        }

        public static void StartWaitProcess(string exePath, string processName, int timeoutSeconds = 300, bool outputProcesses = false)
        {
            Process.Start(exePath);
            WaitProcessStart(processName, timeoutSeconds, outputProcesses: outputProcesses);
        }

        public static void MessageBox(string message, bool showDebugMessages = false)
        {
            if (showDebugMessages) Trace.WriteLine("Pausing execution to display a message box");
            System.Windows.Forms.MessageBox.Show(message);
            if (showDebugMessages) Trace.WriteLine("Continuing execution");
        }

        #endregion

        #region Serialization

        public static string Json(dynamic obj)
        {
            return JsonConvert.SerializeObject(obj, Formatting.Indented);
        }

        public static dynamic YamlDeserialize(string path)
        {
            return YamlDeserialize<dynamic>(path);
        }

        public static T YamlDeserialize<T>(string path)
        {
            string fullPath = Path.GetFullPath(path);
            string yamlText = File.ReadAllText(fullPath);

            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(PascalCaseNamingConvention.Instance)
                .Build();

            T yamlDeserialized = deserializer.Deserialize<T>(yamlText);

            return yamlDeserialized;
        }

        #endregion

        #region DataPrinting

        public static string PrintDataTable(DataTable dtable)
        {
            StringBuilder info = new StringBuilder("DataTable Pretty Print:\n");

            for (int i = 0; i < dtable.Rows.Count; i++)
            {
                DataRow row = dtable.Rows[i];
                info.AppendLine($"  DataRow [{i}]:");

                foreach (DataColumn col in dtable.Columns)
                {
                    info.AppendLine($"    {col.ColumnName}: \"{row[col]}\"");
                }

                info.AppendLine();
            }

            return info.ToString();
        }

        public static string DataRowToString(DataRow row)
        {
            if (row == null) return null;

            StringBuilder dataRow = new StringBuilder();

            foreach (DataColumn col in row.Table.Columns)
            {
                dataRow.Append($"{col}={row[col.ColumnName]}; ");
            }

            return dataRow.ToString();
        }

        #endregion

        #region DataConvertions

        /// <summary>
        /// Converts a COMCollection, where it's impossible to use LINQ, to a ICollection which is LINQable.
        /// </summary>
        /// <typeparam name="T">Collection members type</typeparam>
        /// <param name="comObj">COMCollection object</param>
        /// <returns>ICollection from input COMCollection</returns>
        public static ICollection<T> COMCollectionToICollection<T>(dynamic comObj)
        {
            ICollection<T> collection = new List<T>();

            foreach (var item in comObj)
                collection.Add(item);

            return collection;
        }

        #endregion

        #region CmdPrompt

        public static void RunAsAdmin(object fileName)
        {
            RunAsAdmin(fileName as string);
        }

        public static void RunAsAdmin(string fileName)
        {
            Process proc = new Process();
            proc.StartInfo.FileName = GetFullPath(fileName);
            proc.StartInfo.UseShellExecute = true;
            proc.StartInfo.Verb = "runas";
            proc.Start();
        }

        public static CmdOutput RunPromptCommand(string cmd, string arguments = null, bool redirectErr = true, bool asAdmin = false, 
            Dictionary<string,string> environment = null)
        {

            ProcessStartInfo startOptions = new ProcessStartInfo();
            //startOptions.FileName = "reg";
            //startOptions.Arguments = @"query ""HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon"" /v version";
            startOptions.FileName = cmd;
            startOptions.Arguments = arguments;

            startOptions.CreateNoWindow = true;

            startOptions.UseShellExecute = false; // must be false to be able to redirect output below and change env
            startOptions.RedirectStandardOutput = true;
            startOptions.RedirectStandardError = redirectErr;

            environment?.ToList().ForEach(x => startOptions.Environment[x.Key] = x.Value);

            if (asAdmin)
                startOptions.Verb = "runas";

            Process ps = Process.Start(startOptions);

            string cmdStdOut = ps.StandardOutput.ReadToEnd();
            string cmdStdErr = ps.StandardError.ReadToEnd();
            ps.WaitForExit();

            //Console.WriteLine($"The command output was:\n{cmdOutput}");

            return new CmdOutput(cmdStdOut, cmdStdErr);
        }

        #endregion

        #region APIs

        // returns true if file was downloaded, otherwise returns false
        public static void DownloadFile(string url, string downloadPath)
        {
            downloadPath = GetFullPath(downloadPath);

            using (WebClient client = new WebClient())
            {
                client.DownloadFile(url, downloadPath);
            }

            if (! File.Exists(downloadPath))
                throw new RpaLibFileNotFoundException(downloadPath);
        }

        public static string MakeApiCall(string url, string saveAs = null)
        {
            Task<string> makeApiCallAsync = MakeApiCallAsync(url, saveAs);
            Task.WaitAll(makeApiCallAsync);

            return makeApiCallAsync.Result;
        }

        public static async Task<string> MakeApiCallAsync(string url, string saveAs = null)
        {
            string content;

            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(url);
                content = await response.Content.ReadAsStringAsync();

                if (saveAs != null)
                {
                    string fullFilePath = Ut.GetFullPath(saveAs);
                    FileStream fs = File.OpenWrite(fullFilePath);
                    await response.Content.CopyToAsync(fs);
                    fs.Close();
                }
            }
            
            return content;
        }

        public static CmdOutput Curl (string url, string request = "GET", string[] header = null, string data = null, string uploadFilePath = null)
        {
            // impossible with certificate, he always look inside /etc/ for the cert, but not runs in a full cyg environment
            //string certPath = Path.Combine(CURLPATH, "ca-bundle.crt");
            string commonArgs = $"-v -i -k";
            var args = new StringBuilder($"{commonArgs} --request {request.ToUpper()} --url '{url}'");

            if (header != null)
            {
                foreach (var head in header)
                {
                    args.Append($" --header '{head}'");
                }
            }

            if (data != null)
            {
                args.Append($" --data '{data}'");
            }

            if (uploadFilePath != null)
            {
                args = new StringBuilder($"{commonArgs} --upload-file '{uploadFilePath}' '{url}'");
            }

            return Curl(args.ToString());
        }
        
        public static CmdOutput Curl (string arguments)
        {
            string curl = Path.Combine(CURLPATH, "curl.exe");
            var output = RunPromptCommand(curl, arguments, environment: CYGPATH);

            Trace.WriteLine(output, color: ConsoleColor.Yellow);

            return output;
        }

        public static HttpResponseMessage CurlUploadFile (string filePath, string url)
        {
            return CurlSharp("PUT", url, uploadFilePath: filePath);
        }

        public static HttpResponseMessage CurlSharp(string request, string url, string[] header = null, string data = null, string uploadFilePath = null)
        {
            Task<HttpResponseMessage> curlTask = CurlAsync(url, header, data, uploadFilePath, request);
            Task.WaitAll(curlTask);

            if ((int)curlTask.Result.StatusCode == 200)
            {
                
            }

            var content = curlTask.Result.Content.ReadAsStringAsync();
            Task.WaitAll(content);
            Trace.WriteLine(string.Join("\n",
                $"Status Code: {(int)content.Status}",
                $"Body:",
                $"{content.Result}"), color: ConsoleColor.Yellow);

            return curlTask.Result;
        }

        public static async Task<HttpResponseMessage> CurlAsync(string url, string[] header = null, string data = "BINARY_DATA", string uploadFile = null, string request = "PUT")
        {
            HttpResponseMessage response;

            using (var httpClient = new HttpClient())
            {
                using (var req = new HttpRequestMessage(new HttpMethod(request.ToUpper()), "https://pipefy-production.s3-sa-east-1.amazonaws.com/orgs/ce63-a26b-412f-9d27-3a5776e16e/uploads/45da74d3-0e92-4e1c-a04e-2acc97169a3/SampleFile.pdf?...Signature=fa76e8bf28f88d8ceec1df8219912e103ad15f5ab4668f0fc9cea69109991aa"))
                {
                    if (data != null)
                    {
                        req.Content = new StringContent(data);
                    }
                    else if (uploadFile != null)
                    {
                        req.Content = new ByteArrayContent(File.ReadAllBytes(uploadFile));
                    }

                    if (header != null)
                    {
                        foreach (var head in header)
                        {
                            if (IsMatch(head, @"^Content-Type:\s*"))
                            {
                                req.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(Ut.Replace(head, @"^Content-Type:\s*", string.Empty));
                            }
                        }
                    }

                    response = await httpClient.SendAsync(req);
                    //response.EnsureSuccessStatusCode();
                }
            }

            Trace.WriteLine(response.StatusCode);

            return response;
        }

        #endregion

        #region Aliases

        public static int[] Range(int start, int count)
        {
            return Enumerable.Range(start, count).ToArray();
        }

        #endregion

        #region GUI

        public static void PopUp(string message)
        {
            // pause execution and display a pop-up, continue execution after user click ok
            FormsMsgBox.Show(message);

            Trace.WriteLine(message);
        }

        public static DialogResult PopUpQuestion(string question,
            string ifYes = "Pressed \"Yes\".", string ifNo = "Pressed \"No\".", string ifCancel = "Message box cancelled")
        {
            DialogResult dr = FormsMsgBox.Show(question, "Log Yes/No Question Pop-Up", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes) Trace.WriteLine(ifYes);
            else if (dr == DialogResult.No) Trace.WriteLine(ifNo);
            else Trace.WriteLine(ifCancel);

            return dr;
        }

        #endregion
    }
}
