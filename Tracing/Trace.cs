﻿using System;
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
            LogFile = File.Open(Rpa.GetFullPath(filePath), FileMode.Append);
            var textWritterTraceListener = new TextWriterTraceListener(LogFile);
            textWritterTraceListener.Name = "traceToFile";
            SysTrace.Listeners.Add(textWritterTraceListener);
        }

        public static void WriteLine(string message, bool withTimeSpec = true, bool breakLineBeforeP = true)
        {
            string msg = withTimeSpec ? DateTime.Now.ToString(@"[HH:mm:ss.fff] ") + message : message;
            msg = breakLineBeforeP ? "\n" + msg : msg;
            SysTrace.WriteLine(msg);
        }

        public static void WriteLine(object value, bool withTimeSpec = true)
        {
            WriteLine(value.ToString(), withTimeSpec);
        }

        public static void WriteRedLine(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            SysTrace.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;
        }

        public static void CloseLogFile()
        {
            LogFile.Close();
        }

    }
}