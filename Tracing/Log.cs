using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace RpaLib.Tracing.Legacy
{
    public class Log
    {
        private static string LogTime
        {
            get
            {
                return DateTime.Now.ToString("[dd/MM/yyyy HH:mm:ss.fff]");
            }
        }

        private static string PathDateTime
        {
            get
            {
                return DateTime.Now.ToString("_dd-MM-yyyy_HHmmss");
            }
        }
        public Log(string logPath, LogType logType = LogType.JustConsole, bool appendDateTimeToPath = true)
        {
            if (appendDateTimeToPath)
            {
                string fileName = Path.GetFileNameWithoutExtension(logPath);
                string extension = Path.GetExtension(logPath);
                string basePath = Path.GetDirectoryName(logPath);

                logPath = Path.Combine(basePath, fileName + PathDateTime + extension);
            }

            switch (logType)
            {
                case LogType.ConsoleAndFile:
                    Trace.TraceToConsoleAndFile(logPath);
                    break;

                case LogType.JustConsole:
                    Trace.CopyTraceToConsole();
                    break;

                case LogType.JustFile:
                    Trace.CopyTraceToFile(logPath);
                    break;
            }
        }

        public void Write(string logMessage)
        {
            Trace.WriteLine($"{LogTime} {logMessage} {Environment.NewLine}");
        }

        public void Write(object value)
        {
            Trace.WriteLine($"{LogTime} {value} {Environment.NewLine}");
        }

        public void Dispose ()
        {
            Trace.CloseLogFile();
        }

        ~Log()
        {
            Dispose();
        }

    }
}
