using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class CardField
    {
        public List<string> ArrayValue { get; set; }
        public List<User> AssigneeValues { get; set; }
        //public List<PublicRepoItemTypes> ConnectedRepoItems { get; set; }
        public DateTime Date_Value { get; set; }
        public DateTime DatetimeValue { get; set; }
        //public MinimalField Field { get; set; }
        public DateTime FilledAt { get; set; }
        public float FloatValue { get; set; }
        public string IndexName { get; set; }
        public List<FieldLabel> LabelValues { get; set; }
        public string Name { get; set; }
        public string NativeValues { get; set; }
        public PhaseField PhaseField { get; set; }
        public string ReportValue { get; set; }
        public DateTime UpdatedAt { get; set; }
        public string Value { get; set; }

        /*
         * Deprecated:
         *  - ConnectedRepoItems
         */
    }
}