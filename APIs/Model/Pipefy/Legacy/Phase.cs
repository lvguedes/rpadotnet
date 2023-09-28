//using DocuSign.eSign.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class Phase : IPipefyObject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public PhaseFields Fields { get; set; } = new PhaseFields();
        public Cards Cards { get; set; } = new Cards();

        public object this[string id]
        {
            get => AF.FindFieldValueById(id, Fields.FieldsList.ToArray());
            set => AF.AssignValueToIdInArray(id, Fields.FieldsList.ToArray(), value);
        }
    }
}
