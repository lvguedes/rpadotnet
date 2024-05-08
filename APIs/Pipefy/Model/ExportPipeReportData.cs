using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class ExportPipeReportData
    {
        public string ClientMutationId { get; set; }
        public PipeReportExport PipeReportExport { get; set; }
    }
}
