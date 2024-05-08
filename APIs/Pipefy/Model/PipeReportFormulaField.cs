using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class PipeReportFormulaField
    {
        public double Avg { get; set; }
        public double Count { get; set; }
        public string IndexName { get; set; }
        public string Label { get; set; }
        public double Max { get; set; }
        public double Min { get; set; }
        public string SelectedFormula { get; set; }
        public double Sum { get; set; }
    }
}
