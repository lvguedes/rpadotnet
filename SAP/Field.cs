using sapfewse;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Tracing;

namespace RpaLib.SAP
{
    public class Field : SapComponent
    {
        public Session Session { get; private set; }
        public Type Datatype { get; set; }
        public string Text
        {
            get => GetText();
            set => SetText(value);
        }

        public Field(Session session) : base(session)
        {
            Session = session;
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
            Trace.WriteLine($"Starting to try to extract field {Name} text. ({FullPathId})");

            string text = Session.FindById<GuiVComponent>(fullPathId).Text;

            Trace.WriteLine(string.Join(Environment.NewLine,
                $"Captured field text from some Tab:",
                $"    Field: {Name}",
                $"    Value: {text}"));

            return text;
        }

        public void SetText(string value) => SetText(FullPathId, value);
        public void SetText(string fullPathId, string value)
        {
            Session.FindById<GuiVComponent>(fullPathId).Text = value;
            Trace.WriteLine($"Field \"{Name}\" text changed. New value: \"{value}\"");
        }

        public void Focus() => Focus(FullPathId);
        public void Focus(string fullPathId)
        {
            Trace.WriteLine($"Trying to move focus to field \"{Name}\"");
            Session.FindById<GuiVComponent>(fullPathId).SetFocus();
            Session.FindById<GuiTextField>(fullPathId).CaretPosition = 0;
            Trace.WriteLine($"Focus set to field: \"{Name}\". Carret position: 0.");
        }
    }
}
