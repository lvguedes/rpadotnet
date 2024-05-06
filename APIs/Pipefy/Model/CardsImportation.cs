using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class CardsImportation
    {
        public DateTime? CreatedAt { get; set; }
        public User CreatedBy { get; set; }
        public string DateFormatted { get; set; }
        public string Id { get; set; }
        public int ImportedCards { get; set; }
        public string Status { get; set; }
        public string Url { get; set; }
    }
}
