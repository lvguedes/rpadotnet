using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Exception
{
    public class CardFieldNotFoundException : System.Exception
    {
        public CardFieldNotFoundException(string cardId, string fieldName, int totalFieldsFound) : base(GetMsg(cardId, fieldName, totalFieldsFound)) { }

        public static string GetMsg(string cardId, string fieldName, int totalFieldsFound)
        {
            return $"The field with name \"{fieldName}\" wasn't found in the card ID \"{cardId}\". (Number of fields found: \"{totalFieldsFound}\").";
        }
    }
}
