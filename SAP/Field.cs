using sapfewse;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class Field : SapComponent<GuiTextField>
    {
        public string Xml { get; set; }
        public string Sap { get; set; }
        public Type Datatype { get; set; }
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        public dynamic ConvertToDatatype(string value)
        {
            if (Datatype.Equals(typeof(string)))
                return value;
            else if (Datatype.Equals(typeof(double)))
                return double.Parse(value, new CultureInfo("pt-BR"));
            else if (Datatype.Equals(typeof(long)))
                return long.Parse(value);
            else if (Datatype.Equals(typeof(int)))
                return int.Parse(value);
            else if (Datatype.Equals(typeof(DateTime)))
                return DateTime.ParseExact(value, "dd.MM.yyyy", null);
            else
                throw new FieldDatatypeConversionException();
        }

        public string GetText() => GetText(FullPathId);
        public string GetText(string fullPathId)
        {
            Tracing.Log.Write($"Starting to try to extract field {Name} text. ({FullPathId})");

            string text = (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).Text;

            Tracing.Log.Write(string.Join(Environment.NewLine,
                $"Captured field text from some Tab:",
                $"    Field: {Name}",
                $"    Value: {text}"));

            return text;
        }

        public void SetText(string value) => SetText(FullPathId, value);
        public void SetText(string fullPathId, string value)
        {
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).Text = value;
            Tracing.Log.Write($"Field \"{Name}\" text changed. New value: \"{value}\"");
        }

        public void Focus() => Focus(FullPathId);
        public void Focus(string fullPathId)
        {
            Tracing.Log.Write($"Trying to move focus to field \"{Name}\"");
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiVComponent).SetFocus();
            (RpaLib.SAP.Sap.Session.FindById(fullPathId) as GuiTextField).CaretPosition = 0;
            Tracing.Log.Write($"Focus set to field: \"{Name}\". Carret position: 0.");
        }
    }
}
