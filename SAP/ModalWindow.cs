using sapfewse;
using System;
using System.Collections.Generic;
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
        public GuiModalWindow GuiModalWindow { get; private set; }

        public Session Session { get; private set; }

        public ModalWindow(string pathId, Session session)
        {
            GuiModalWindow = session.FindById<GuiModalWindow>(pathId);
            Session = session;
        }

        public T FindById<T>(string pathId) => (T)GuiModalWindow.FindById(pathId);

        public T[] FindByType<T>(string typeName, bool showFound = false) => Sap.FindByType<T>(GuiModalWindow, typeName, showFound);

        public T[] FindByType<T>(bool showFound = false) => Sap.FindByType<T>(GuiModalWindow, showFound);
    }
}
