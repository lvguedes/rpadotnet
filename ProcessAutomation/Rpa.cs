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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Data;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization;
using System.Net.Http;
using Newtonsoft.Json;

/*
Reference Assemblies:
  - System.IO.Compression.FileSystem
*/


/// Generic RPA functions
namespace RpaLib.ProcessAutomation
{
    public delegate void VoidFuncVoid();
    public static class Rpa
    {
        [DllImport("Kernel32.dll", SetLastError = true)]
        public static extern int SetStdHandle(int device, IntPtr handle);

        #region RegularExpressions

        public static string Replace(string input, string pattern, string replacement,
            RegexOptions regexOptions = RegexOptions.IgnoreCase) => Regex.Replace(input, pattern, replacement);
        public static string Match(string input, string pattern, RegexOptions regexOptions = RegexOptions.IgnoreCase) => Regex.Match(input, pattern, regexOptions).Value;

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

            throw new FileNotFoundException();
        }

        public static bool IsFileInDir(string fileNameRegex, string dirPath)
        {
            string dirFullPath = GetFullPath(dirPath);

            foreach (string file in Directory.GetFiles(dirFullPath))
            {
                if (Rpa.IsMatch(file, fileNameRegex))
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
            var fullFilePath = Rpa.GetFullPath(filePath);

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
                        throw new TimeoutException();
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

        public static string RunPromptCommand(string cmd, string arguments = null)
        {

            ProcessStartInfo startOptions = new ProcessStartInfo();
            //startOptions.FileName = "reg";
            //startOptions.Arguments = @"query ""HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon"" /v version";
            startOptions.FileName = cmd;
            startOptions.Arguments = arguments;

            startOptions.CreateNoWindow = true;

            startOptions.UseShellExecute = false; // must be false to be able to redirect output below
            startOptions.RedirectStandardOutput = true;

            startOptions.Verb = "runas";

            Process ps = Process.Start(startOptions);

            string cmdOutput = ps.StandardOutput.ReadToEnd();
            ps.WaitForExit();

            //Console.WriteLine($"The command output was:\n{cmdOutput}");

            return cmdOutput;
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
                throw new FileNotFoundException(downloadPath);
        }

        public static string MakeApiCall(string url, string saveAs = null)
        {
            Task<string> makeApiCallAsync = MakeApiCallAsync(url, saveAs);
            Task.WaitAll(makeApiCallAsync);

            return makeApiCallAsync.Result;
        }

        public static async Task<string> MakeApiCallAsync(string url, string saveAs = null)
        {
            var client = new HttpClient();
            var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();

            if (saveAs != null)
            {
                string fullFilePath = Rpa.GetFullPath(saveAs);
                FileStream fs = File.OpenWrite(fullFilePath);
                await response.Content.CopyToAsync(fs);
                fs.Close();
            }

            return content;
        }

        #endregion
    }
}
