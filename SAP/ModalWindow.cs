using sapfewse;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    /// <summary>
    /// Class that represent a Pop-Up window (GuiModalWindow) within SAP
    /// </summary>
    public class ModalWindow : SapComponent<GuiModalWindow>
    {
        public SapComWrapper<GuiModalWindow> GuiModalWindow
        {
            get => Session.FindById<GuiModalWindow>(PathId);
        }

        public ModalWindow(Session session, string pathId) : base(session, pathId)
        { }

        public SapComWrapper<T> FindById<T>(string pathId) => GuiModalWindow.FindById<T>(pathId);

        public SapComWrapper<T>[] FindByType<T>(string typeName, bool showFound = false) => Sap.FindByType<T>((GuiComponent)GuiModalWindow.Com, typeName, showFound);

        public SapComWrapper<T>[] FindByType<T>(bool showFound = false) => Sap.FindByType<T>((GuiComponent)GuiModalWindow.Com, showFound);
    }
}
