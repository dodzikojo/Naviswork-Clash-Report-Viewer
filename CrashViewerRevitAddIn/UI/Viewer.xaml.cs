using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CrashViewerRevitAddIn.UI
{
    /// <summary>
    /// Interaction logic for Viewer.xaml
    /// </summary>
    public partial class Viewer : Page, IDockablePaneProvider
    {
        // fields
        public ExternalCommandData eData = null;
        public Document doc = null;
        public UIDocument uidoc = null;

        public Viewer()
        {
            InitializeComponent();
        }

        //public void CustomInitiator(ExternalCommandData e)
        //{
        //    // ExternalCommandData and Doc
        //    eData = e;
        //    doc = e.Application.ActiveUIDocument.Document;
        //    uidoc = eData.Application.ActiveUIDocument;

        //    // get the current document name
        //    docName.Text = doc.PathName.ToString().Split('\\').Last();
        //    // get the active view name
        //    viewName.Text = doc.ActiveView.Name;
        //    // call the treeview display method
        //    //DisplayTreeViewItem();
        //}

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            // wpf object with pane's interface
            data.FrameworkElement = this as FrameworkElement;
            // initial state position
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.PropertiesPalette
            };

        }
    }
}
