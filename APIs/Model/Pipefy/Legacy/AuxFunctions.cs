using Newtonsoft.Json.Serialization;
using RpaLib.SAP;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy.Legacy
{
    public class AuxFunctions
    {
        private static bool FieldIdExistsP<T>(string id, T[] fields) where T : IPipefyObject, IPipefyField
        {
            try
            {
                FindFieldValueById(id, fields);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        public static IPipefyObject FindFieldValueById(string id, List<IPipefyObject> fields)
        {
            return FindFieldValueById(id, fields.ToArray());
        }
        public static T FindFieldValueById<T>(string id, T[] fields) where T : IPipefyObject
        {
            foreach (var field in fields)
            {
                if (field.Id == id)
                {
                    return field;
                }
            }

            throw new ArgumentOutOfRangeException(
                nameof(id),
                $"The ID \"{id}\" was not found within phase fields");
        }

        public static void AssignValueToIdInArray<T>(string id, List<T> pipefyObjects, object value)
        {
            //AssignValueToIdInArray(id, pipefyObjects.ToArray(), value);
        }
        public static void AssignValueToIdInArray<T>(string id, T[] pipefyObjects, object value) where T : IPipefyField
        {
            if (FieldIdExistsP(id, pipefyObjects))
            {
                pipefyObjects.Cast<T>().Where(x => x.Id == id).FirstOrDefault().Value = value;
                /*
                var queriedField = from field in Fields
                                   where field.Id == id
                                   select field;
                queriedField.FirstOrDefault().Value = value;
                */
            }
            else
            {
                var fieldsList = pipefyObjects.ToList();
                fieldsList.Add((T)value);
                pipefyObjects = fieldsList.ToArray();
            }
        }
    }
}
