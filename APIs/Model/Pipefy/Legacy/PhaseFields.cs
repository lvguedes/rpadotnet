using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class PhaseFields
    {
        public List<PhaseField> FieldsList { get; set; } = new List<PhaseField>();

        public PhaseField this[string id]
        {
            get => (PhaseField)AF.FindFieldValueById(id, FieldsList.ToArray());
            set => AF.AssignValueToIdInArray(id, FieldsList.ToArray(), value);
        }

        public PhaseFields() { }
        public PhaseFields(params PhaseField[] phaseFields)
            : this(phaseFields.ToList()){ }
        public PhaseFields(List<PhaseField> phaseField)
        {
            FieldsList = phaseField;
        }

        public void Add(params PhaseField[] phaseFields)
        {
            foreach (var phaseField in phaseFields)
            {
                FieldsList.Add(phaseField);
            }
        }

        public void Add(List<PhaseField> phaseFields)
        {
            Add(phaseFields.ToArray());
        }

        public void Add(PhaseFields phaseFields)
        {
            Add(phaseFields.FieldsList);
        }



        public IEnumerator GetEnumerator() => FieldsList.GetEnumerator();
    }
}
