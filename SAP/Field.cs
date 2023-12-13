using sapfewse;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Tracing;
using RpaLib.SAP.Exceptions;

namespace RpaLib.SAP
{
    public class Field : SapComponent<GuiCTextField>
    {
        public Type Datatype { get; set; }
        //public string Text
        //{
        //    get => GetText();
        //    set => SetText(value);
        //}

        public Field(Session session, string pathId)
            : base (session, pathId)
        { }

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

        //public string GetText() => GetText(PathId);
        //public string GetText(string fullPathId)
        //{
        //    Trace.WriteLine($"Starting to try to extract field {Name} text. ({PathId})");

        //    string text = Session.FindById<GuiVComponent>(fullPathId).Text;

        //    Trace.WriteLine(string.Join(Environment.NewLine,
        //        $"Captured field text from some Tab:",
        //        $"    Field: {Name}",
        //        $"    Value: {text}"));

        //    return text;
        //}

        //public void SetText(string value) => SetText(PathId, value);
        //public void SetText(string fullPathId, string value)
        //{
        //    Session.FindById<GuiVComponent>(fullPathId).Text = value;
        //    Trace.WriteLine($"Field \"{Name}\" text changed. New value: \"{value}\"");
        //}

        public void Focus() => Focus(PathId);
        public void Focus(string fullPathId)
        {
            Trace.WriteLine($"Trying to move focus to field \"{Com.Id}\"");
            Session.FindById<GuiVComponent>(fullPathId).SetFocus();
            Session.FindById<GuiTextField>(fullPathId).Com.CaretPosition = 0;
            Trace.WriteLine($"Focus set to field: \"{Com.Id}\". Carret position: 0.");
        }

        public void SelectPossibleValue(string pathIdTable, string value, int[] dropLines = null, int header = 0)
        {
            Sap.SelectPossibleValue(Session, this, pathIdTable, value, header, dropLines);
        }

        public void OpenPossibleValues() => Sap.OpenPossibleValues(Session, this);

        public LabelTable OpenPossibleValues(string guiUserAreaId)
        {
            return Sap.OpenPossibleValues(Session, this, guiUserAreaId);
        }

        public void ClosePossibleValues() => Session.PressEsc();
    }
}
