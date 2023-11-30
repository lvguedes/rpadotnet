using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using sapfewse;
using RpaLib.ProcessAutomation;

namespace RpaLib.SAP
{
    public abstract class SapComponent<T> : SapComWrapper<T>
    {
        private string _pathId;
        public Session Session { get; protected set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string PathId
        {
            get => _pathId;
            //set
            //{
            //    Tracing.Log.Write(string.Join(Environment.NewLine,
            //        $"Setting the PathID of wrapper to [{To<GuiComponent>().Type}]{To<GuiComponent>().Id} \"{To<GuiVComponent>()?.Text}\" (nickname: \"{Name}\"):",
            //        $"  Old value: \"{_pathId}\"",
            //        $"  New value: \"{value}\""));
            //    _pathId = value;
            //}
        }

        public SapComponent(Session session, string pathId)
            : this(session, session.FindById<T>(pathId).Com)
        {
            _pathId = pathId;
        }

        public SapComponent(Session session, T component)
            : base(component)
        {
            Session = session;
            _pathId = _pathId ?? Ut.Replace(((GuiComponent)component).Id, @"/app/con\[\d+\]/ses\[\d+\]/", string.Empty);
        }

        //public dynamic FindById(string fullPathId) => Session.FindById(fullPathId);
        //public U FindById<U>(string fullPathId) => Session.FindById<U>(fullPathId);
    }
}
