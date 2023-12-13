using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.ProcessAutomation
{
    public class CmdOutput
    {
        public string StdOut { get; private set; }
        public string StdErr { get; private set; }

        public CmdOutput(StreamReader stdout, StreamReader stderr = null) 
            : this(stdout.ReadToEnd(), stderr.ReadToEnd()){ }

        public CmdOutput(string stdout, string stderr = null)
        {
            StdOut = stdout;
            StdErr = stderr;
        }

        public override string ToString()
        {
            var indent = "    ";

            Func<string, string> indentOutput = 
                (string t) => string.Join(Environment.NewLine + indent, t.Replace(Environment.NewLine, "\n").Split('\n'));

            return string.Join(Environment.NewLine,
                $"Standard Output:",
                indentOutput(StdOut),
                string.Empty,
                $"Standard Error:",
                indentOutput(StdErr));
        }
    }
}
