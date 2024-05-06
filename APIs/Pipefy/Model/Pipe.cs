using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class Pipe
    {
        public List<Phase> Phases { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }
        public string Noun { get; set; }
        public List<Label> Labels { get; set; }
        public string OrganizationId { get; set; }
        public Organization Organization { get; set; }
        public List<PhaseField> StartFormFields { get; set; }
    }
}
