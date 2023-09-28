using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class Phases : IEnumerable
    {
        public List<Phase> PhasesList { get; set; }

        //public Phase this[string id]
        //{
        //    get => (Phase)AF.FindFieldValueById(id, PhasesList.ToArray());
        //    set => AF.AssignValueToIdInArray(id, PhasesList, value);
        //}

        public Phases() { }
        public Phases(Phase[] phases)
            : this(phases.ToList()){ }
        public Phases(List<Phase> phases)
        {
            PhasesList = phases;
        }

        public IEnumerator GetEnumerator() => PhasesList.GetEnumerator();
       
    }
}
