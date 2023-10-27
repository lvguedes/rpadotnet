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

        private static TraceListener[] SearchListener(string listenerName)
        {
            var foundListeners = SysTrace.Listeners.Cast<TraceListener>().Where(x => x.Name == listenerName).ToArray();
            return foundListeners;
        }

        // Copy the Trace stream to stdOut or stdErr
        public static void CopyTraceToConsole(bool toStdErr = false)
        {
            const string consoleTracerName = "traceToConsole";

            // Add the listener only if it's not yet added
            if (SearchListener(consoleTracerName).Length == 0)
            {
                var consoleTracer = new ConsoleTraceListener(toStdErr);

                consoleTracer.Name = consoleTracerName;

                SysTrace.Listeners.Add(consoleTracer);
            }
            else
            {
                Trace.WriteLine($"Trace listener \"{consoleTracerName}\" already exists", color: ConsoleColor.Yellow);
            }
        }

        // Copy the Trace stream to a file
        public static void CopyTraceToFile(string filePath)
        {
            const string fileTracerName = "traceToFile";

            if (SearchListener(fileTracerName).Length == 0)
            {
                var fullFilePath = Rpa.GetFullPath(filePath);

                //Rpa.CreateFileIfNotExists(fullFilePath);

                LogFile = File.Open(fullFilePath, FileMode.Append);

                var textWritterTraceListener = new TextWriterTraceListener(LogFile);
                textWritterTraceListener.Name = "traceToFile";
                SysTrace.Listeners.Add(textWritterTraceListener);
            }
            else
            {
                Trace.WriteLine($"Trace listener \"{fileTracerName}\" already exists", color: ConsoleColor.Yellow);
            }
        }

        public static void WriteLine(string message, bool withTimeSpec = true, bool breakLineBeforeP = true, ConsoleColor color = ConsoleColor.White)
        {
            string msg = withTimeSpec ? DateTime.Now.ToString(@"[HH:mm:ss.fff] ") + message : message;
            msg = breakLineBeforeP ? "\n" + msg : msg;
            WriteColorLine(msg, color);
        }

        public static void WriteLine(object value, bool withTimeSpec = true, bool breakLineBeforeP = true, ConsoleColor color = ConsoleColor.White)
        {
            WriteLine(value.ToString(), withTimeSpec, breakLineBeforeP, color);
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
