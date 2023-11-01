using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class MinimalField
    {
        public string Id { get; set; }
        public List<string> Options { get; set; }
        public string Type { get; set; }
        public string Description { get; set; }
        public string Help { get; set; }
        public float Index { get; set; }
        public string IndexName { get; set; }
        public string InternalId { get; set; }
        public string Label { get; set; }
        public string Uuid { get; set; }
    }
}
