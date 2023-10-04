using AutoIt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SDKs.AutoIt
{
    public enum WindowState
    {
        Normal = AutoItX.SW_NORMAL,
        Minimize = AutoItX.SW_MINIMIZE,
        Maximize = AutoItX.SW_MAXIMIZE,
        Hide = AutoItX.SW_HIDE,
        Show = AutoItX.SW_SHOW,
        ShowDefault = AutoItX.SW_SHOWDEFAULT,
        ShowNormal = AutoItX.SW_SHOWNORMAL,
        ShowMaximized = AutoItX.SW_SHOWMAXIMIZED,
        ShowMinimized = AutoItX.SW_SHOWMINIMIZED,
        ShowMinNoActive = AutoItX.SW_SHOWMINNOACTIVE,
        ShowNoActivate = AutoItX.SW_SHOWNOACTIVATE,
        ShowNA = AutoItX.SW_SHOWNA,
        ForceMinimize = AutoItX.SW_FORCEMINIMIZE,
    }
}
