using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class Card : IPipefyObject
    {
        public string Id { get; set; }
        public CardFields Fields { get; set; } = new CardFields();

        //public object this[string id]
        //{
        //    get => AF.FindFieldValueById(id, Fields.FieldsList.ToArray());
        //    set => AF.AssignValueToIdInArray(id, Fields.FieldsList.ToArray(), value);
        //}
    }
}
