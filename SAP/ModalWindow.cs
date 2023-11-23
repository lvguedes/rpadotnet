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
    public class ModalWindow
    {
        public string PathId { get; }

        public GuiModalWindow GuiModalWindow
        {
            get => Session.FindById<GuiModalWindow>(PathId);
        }

        public Session Session { get; private set; }

        public ModalWindow(string pathId, Session session)
        {
            PathId = pathId;
            Session = session;
        }

        public T FindById<T>(string pathId) => (T)GuiModalWindow.FindById(pathId);

        public T[] FindByType<T>(string typeName, bool showFound = false) => Sap.FindByType<T>(GuiModalWindow, typeName, showFound);

        public T[] FindByType<T>(bool showFound = false) => Sap.FindByType<T>(GuiModalWindow, showFound);
    }
}
