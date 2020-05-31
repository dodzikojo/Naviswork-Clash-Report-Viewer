using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrashViewerRevitAddIn
{
    class Globals
    {
        public const string ApplicationName = "Crash Viewer";
        public const string DiagnosticsTabName = "DockableDialogs";
        public const string DiagnosticsPanelName = "DockableDialogs Panel";

        public const string RegisterPage = "Register Page";
        public const string ShowPage = "Show Page";
        public const string HidePage = "Hide Page";


        public static DockablePaneId sm_UserDockablePaneId = new DockablePaneId(new Guid("{3BAFCF52-AC5C-4CF8-B1CB-D0B1D0E90237}"));
    }
}
