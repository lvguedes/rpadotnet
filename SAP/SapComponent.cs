using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SapLegacy = RpaLib.SAP.Legacy.Sap;

namespace RpaLib.SAP
{
    public abstract class SapComponent
    {
        private string _fullPathId;
        public Session Session { get; private set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string FullPathId
        {
            get => _fullPathId;
            set
            {
                Tracing.Log.Write(string.Join(Environment.NewLine,
                    $"Setting the FullPathID of {this.GetType()} \"{Name}\":",
                    $"  Old value: \"{_fullPathId}\"",
                    $"  New value: \"{value}\""));
                _fullPathId = value;
            }
        }
        public string Id { get; set; }
        public static string BasePathId { get; set; }

        public SapComponent(Session session)
        {
            Session = session;
        }

        //public dynamic FindById(string fullPathId) => Session.FindById(fullPathId);
        //public U FindById<U>(string fullPathId) => Session.FindById<U>(fullPathId);
    }
}
