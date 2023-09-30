using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SapLegacy = RpaLib.SAP.Legacy.Sap;

namespace RpaLib.SAP
{
    public abstract class SapComponent<T> //: Sap
    {
        private string _fullPathId;
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

        public SapComponent()
        {
            //UpdateParentParams();
        }

        public void UpdateParentParams()
        {
            SapLegacy.UpdateConnections();
            SapLegacy.MapExistingSessions();
        }

        public static dynamic FindById(string fullPathId) => SapLegacy.Session.FindById(fullPathId);
        public static U FindById<U>(string fullPathId) => (U)SapLegacy.Session.FindById(fullPathId);
    }
}
