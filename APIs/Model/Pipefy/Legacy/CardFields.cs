using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class CardFields : IEnumerable
    {
        public List<CardField> FieldsList { get; set; } = new List<CardField>();

        public CardField this[string id]
        {
            get => (CardField)AF.FindFieldValueById(id, FieldsList.ToArray());
            set => AF.AssignValueToIdInArray(id, FieldsList.ToArray(), value);
        }

        public CardFields() { }
        public CardFields(CardField[] cardFields)
            :this(cardFields.ToList()) { }
        public CardFields(List<CardField> cardFields)
        {
            FieldsList = cardFields;
        }
        public void Add(params CardField[] cardFields)
        {
            foreach (var cardField in cardFields)
            {
                FieldsList.Add(cardField);
            }
        }

        public void Add(List<CardField> cardFields)
        {
            Add(cardFields.ToArray());
        }

        public void Add(CardFields cardFields)
        {
            Add(cardFields.FieldsList);
        }


        public IEnumerator GetEnumerator() => FieldsList.GetEnumerator();
    }
}
