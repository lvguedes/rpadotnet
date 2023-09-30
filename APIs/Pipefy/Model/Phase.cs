using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    using Id = String;
    public class Phase
    {
        public bool CanReceiveCardDirectlyFromDraft { get; set; }
        public CardConnection Cards { get; set; }
        public List<Phase> CardsCanBeMovedToPhases { get; set; }
        public int CardsCount { get; set; }
        public string Color { get; set; }
        public DateTime CreatedAt { get; set; }
        public string CustomSortingPreferences { get; set; }
        public string Description { get; set; }
        public bool Done { get; set; }
        public int ExpiredCardsCount { get; set; }
        //public List<FieldCondition> FieldConditions { get; set; }
        public List<PhaseField> Fields { get; set; }
        public Id Id { get; set; }
        //public IdentifyTaskEnum IdentifyTask { get; set; }
        public double Index { get; set; }
        public bool IsDraft { get; set; }
        public int LatenessTime { get; set; }
        public string Name { get; set; }
        public List<int> NextPhaseIds { get; set; }
        public List<int> PreviousPhaseIds { get; set; }
        public int RepoId { get; set; }
        public string SequentialId { get; set; }
        public Id Uuid { get; set; }
    }
}
