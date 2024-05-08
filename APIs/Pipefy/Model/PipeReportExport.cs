using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class PipeReportExport
    {
        public string FileURL { get; set; }
        public DateTime? FinishedAt { get; set; }
        public string Id { get; set; }
        public PipeReport Report { get; set; }
        public User RequestedBy { get; set; }
        public DateTime? StartedAt { get; set; }
        public ExpirationState State { get; set; }
    }
}
