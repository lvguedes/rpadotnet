using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class PipeReport
    {
        public int CardCount { get; set; }
        public string Color { get; set; }
        public DateTime? CreatedAt { get; set; }
        public string FeaturedField { get; set; }
        public List<string> Fields { get; set; }
        public string Json { get; set; }
        public string Id { get; set; }
        public DateTime? LastUpdatedAt { get; set; }
        public string Name { get; set; }
        public List<PipeReportFormulaField> SelectedFormulaFields { get; set; }
        public ReportSortDirection SortBy { get; set; }
    }
}
