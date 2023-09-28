using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AF = RpaLib.APIs.Model.Pipefy.Legacy.AuxFunctions;

namespace RpaLib.APIs.Model.Pipefy
{
    public class Cards : IEnumerable
    {
        public List<Card> CardsList { get; set; }

        //public Card this[string id]
        //{
        //    get => (Card)AF.FindFieldValueById(id, CardsList.ToArray());
        //    set => AF.AssignValueToIdInArray(id, CardsList.ToArray(), value);
        //}

        public Cards() { }
        public Cards(Card[] cards)
            : this(cards.ToList()) { }
        public Cards(List<Card> cards)
        {
            CardsList = cards;
        }

        public IEnumerator GetEnumerator() => CardsList.GetEnumerator();
    }
}
