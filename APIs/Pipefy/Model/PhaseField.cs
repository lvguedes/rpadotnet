using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    using Id = String;
    public class PhaseField
    {
        public bool AllChildrenMustBeDoneToFinishParent { get; set; }
        public bool AllChildrenMustBeDoneToMoveParent { get; set; }
        public bool CanConnectExisting { get; set; }
        public bool CanConnectMultiples { get; set; }
        public bool CanCreateNewConnected { get; set; }
        public bool ChildMustExistToFinishParent { get; set; }
        //public PublicRepoUnion ConnectedRepo { get; set; }
        public string CustomValidation { get; set; }
        public string Description { get; set; }
        public bool Editable { get; set; }
        public string Help { get; set; }
        public Id Id { get; set; }
        public double Index { get; set; }
        public string IndexName { get; set; }
        public Id InternalId { get; set; }
        public bool IsMultiple { get; set; }
        public string Label { get; set; }
        public List<string> Options { get; set; }
        public Phase Phase { get; set; }
        public bool Required { get; set; }
        public bool SyncedWithCard { get; set; }
        public String Type { get; set; }
        public Id Uuid { get; set; }

        /*
         * Deprecated
         *  - ConnectedRepo
         */
    }
}
