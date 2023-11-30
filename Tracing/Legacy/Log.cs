using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using RpaLib.ProcessAutomation;
using System.Windows.Forms;  // reference assembly: System.Windows.Forms

namespace RpaLib.Tracing
{
    public static class Log
    {
        private static string _logPath;
        private static string _logName;
        private static string ParentDir { get; set; }
        public static bool DisplayDebugPopups { get; set; } = false;
        public static string LogPath
        {
            set
            {
                ParentDir = Path.GetDirectoryName(value);
                _logName = Path.GetFileNameWithoutExtension(Ut.GetFullPath(value));
                //_logPath = Path.Combine(ParentDir, $"{Date}_{LogName}.txt"); // just the date in the name, make it append
                _logPath = Path.Combine(ParentDir, $"{Date}_{Time}_{LogName}.txt"); // date and time: make it create one file for each exec
                CreateLogPathIfNotExists(_logPath);
            }
            get { return _logPath; }
        }
        public static string LogName
        {
            get
            {
                return _logName;
            }
        }
        public static string Date { get; } = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        public static string Time { get; } = DateTime.Now.ToString("HHmmss", CultureInfo.InvariantCulture);


        public static void CreateLogPathIfNotExists(string filePath = @".\")
        {
            if (!Directory.Exists(ParentDir))
            {
                Directory.CreateDirectory(ParentDir);
            }
            /* Not good because duplicates the log file in the same dir of exe file
            if (!File.Exists(filePath))
            {
                File.Create(filePath).Close();
            }
            */
        }

        public static void Write(string logMessage, string filePath = @".\")
        {
            if (string.IsNullOrEmpty(LogPath)) LogPath = filePath;
            string msg = string.Format("[{0}] {1} {2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"), logMessage, Environment.NewLine);
            Trace.WriteLine(msg);
            //File.AppendAllText(LogPath, msg);
        }

        public static void Write(Exception ex, string filePath = @".\")
        {
            Write(ex.ToString(), filePath);
        }

        public static void WriteBlock(string logMessage, string filePath = @"\.")
        {
            if (string.IsNullOrEmpty(LogPath)) LogPath = filePath;
            string nl = Environment.NewLine;
            string separator = $"#############################################################{nl}";
            string msg = $"{separator}{logMessage}{separator}";
            Write(msg, filePath);
        }

        public static void WriteException(Exception ex, string filePath = @".\")
        {
            if (string.IsNullOrEmpty(LogPath)) LogPath = filePath;
            string msg = string.Join(
                Environment.NewLine,
                $"Exception \"{ex}\" occurred.",
                $"Message: \"{ex.Message}\"",
                $"----------------------",
                $"Stack Trace:",
                $"{ex.StackTrace}",
                $"----------------------",
                $"Source: \"{ex.Source}\"\n",
                $"Inner Exception: \"{ex.InnerException}\"\n",
                $"Data:",
                $"{GetExceptionData(ex.Data)}"
                );

            WriteBlock(msg, filePath);
        }

        private static string GetExceptionData(IDictionary data)
        {
            StringBuilder stringBuilder = new StringBuilder();

            foreach (string key in data.Keys)
            {
                stringBuilder.AppendLine($"    {key}: {data[key]}");
            }

            return stringBuilder.ToString();
        }

        public static void MessageBox(string message)
        {
            // pause execution and display a pop-up, continue execution after user click ok
            if (DisplayDebugPopups)
            {
                System.Windows.Forms.MessageBox.Show(message);
            }
            Write(message);
        }

        public static DialogResult YNQuestion(string question, string ifYes, string ifNo, string ifCancel = "Message box cancelled")
        {
            if (!DisplayDebugPopups) return DialogResult.Abort;

            DialogResult dr = System.Windows.Forms.MessageBox.Show(question, "Log Yes/No Question Pop-Up", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)     Write(ifYes);
            else if (dr == DialogResult.No) Write(ifNo);
            else                            Write(ifCancel);

            return dr;
        }
    }
}

