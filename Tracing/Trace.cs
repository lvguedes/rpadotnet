using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SysTrace = System.Diagnostics.Trace;
using RpaLib.ProcessAutomation;

namespace RpaLib.Tracing
{
    public static class Trace
    {
        private static FileStream LogFile { get; set; }

        public static void TraceToConsoleAndFile(string filePath)
        {
            SysTrace.AutoFlush = true;
            CopyTraceToConsole();
            CopyTraceToFile(filePath);
        }

        // Copy the Trace stream to stdOut or stdErr
        public static void CopyTraceToConsole(bool toStdErr = false)
        {
            var consoleTracer = new ConsoleTraceListener(toStdErr);

            consoleTracer.Name = "traceToConsole";

            SysTrace.Listeners.Add(consoleTracer);
        }

        // Copy the Trace stream to a file
        public static void CopyTraceToFile(string filePath)
        {
            var fullFilePath = Rpa.GetFullPath(filePath);

            //Rpa.CreateFileIfNotExists(fullFilePath);

            LogFile = File.Open(fullFilePath, FileMode.Append);

            var textWritterTraceListener = new TextWriterTraceListener(LogFile);
            textWritterTraceListener.Name = "traceToFile";
            SysTrace.Listeners.Add(textWritterTraceListener);
        }

        public static void WriteLine(string message, bool withTimeSpec = true, bool breakLineBeforeP = true, ConsoleColor color = ConsoleColor.White)
        {
            string msg = withTimeSpec ? DateTime.Now.ToString(@"[HH:mm:ss.fff] ") + message : message;
            msg = breakLineBeforeP ? "\n" + msg : msg;
            WriteColorLine(msg, color);
        }

        public static void WriteLine(object value, bool withTimeSpec = true, ConsoleColor color = ConsoleColor.White)
        {
            WriteLine(value.ToString(), withTimeSpec, color);
        }

        private static void WriteColorLine(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            SysTrace.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White; // turn console color back to default
        }

        public static void CloseLogFile()
        {
            LogFile.Close();
        }

    }
}
