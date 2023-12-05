using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AutoIt;
using RpaLib.Tracing;
using RpaLib.SDKs.AutoIt.Exceptions;
using RpaLib.Exceptions;

namespace RpaLib.SDKs.AutoIt
{
    public class AutoIt
    {
        public const int DefaultTimeout = 5;

        private static void ThrowWinWaitTimeout(string title, string text, int timeout)
        {
            throw new RpaLibTimeoutException($"Window with " +
            $"(title:\"{title}\"; text:\"{text}\"; timeout:{timeout}) " +
                    $"didn't appear within timeout of {timeout} seconds.");
        }
        public static void WinWait(string title, string text = "", int timeout = DefaultTimeout)
        {
            Trace.WriteLine($"AutoIt: WinWait( title: \"{title}\", text: \"{text}\", timeout: {timeout} )");

            var result = AutoItX.WinWait(title, text, timeout);

            if (result != 1)
                ThrowWinWaitTimeout(title, text, timeout);
        }
        public static void WinWaitActive(string title, string text = "", int timeout = DefaultTimeout)
        {

            Trace.WriteLine($"AutoIt: WinWaitActive(title: \"{title}\", text: \"{text}\", timeout: {timeout})");

            var result = AutoItX.WinWaitActive(title, text, timeout);

            if (result != 1)
                ThrowWinWaitTimeout(title, text, timeout);
        }

        public static int Run(string fileName, string workingDir = "", WindowState windowState = WindowState.Normal)
        {
            Trace.WriteLine($"AutoIt: Run(fileName: \"{fileName}\", workingDir: \"{workingDir}\", windowState: {windowState})");
            AutoItX.Run(fileName, workingDir, (int)windowState);

            var error = AutoItX.ErrorCode();

            return error;
        }

        public static void Send(string keys, SendKeysFlag flag = SendKeysFlag.UseSpecialChars)
        {
            Trace.WriteLine($"AutoIt: Send( keys: \"{keys}\", flag: {flag} )");
            AutoItX.Send(keys, (int)flag);
        }

        public static void Sleep(int seconds) => AutoItX.Sleep(1000*seconds);

        public static bool ProcessExists(string nameOrPid)
        {
            var result = AutoItX.ProcessExists(nameOrPid);

            if (result == 0)
                return false;
            else
                return true;
        }
    }
}
