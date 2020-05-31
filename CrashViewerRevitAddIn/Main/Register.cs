using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CrashViewerRevitAddIn.Main
{
    // external command class
    [Transaction(TransactionMode.Manual)]
    public class Register : IExternalCommand
    {
    //    UI.Viewer dockableWindow = null;
    //    ExternalCommandData edata = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //// dockable window
            //UI.Viewer dock = new UI.Viewer();
            //dockableWindow = dock;
            //edata = commandData;


            //// create a new dockable pane id
            //DockablePaneId id = new DockablePaneId(new Guid(Guid.NewGuid().ToString()));
            //try
            //{
            //    // register dockable pane
            //    commandData.Application.RegisterDockablePane(id, "Crash Viewer Dockable",
            //            dockableWindow as IDockablePaneProvider);
            //    TaskDialog.Show("Info Message", "Crash Report Doc has been registered.");
            //    // subscribe document opened event
            //    commandData.Application.Application.DocumentOpened += new EventHandler<Autodesk.Revit.DB.Events.DocumentOpenedEventArgs>(Application_DocumentOpened);
            //    // subscribe view activated event
            //    commandData.Application.ViewActivated += new EventHandler<ViewActivatedEventArgs>(Application_ViewActivated);
            //}
            //catch (Exception ex)
            //{
            //    // show error info dialog
            //    TaskDialog.Show("Info Message", ex.Message);
            //}

            // return result
            return Result.Succeeded;
        }


        // view activated event
        //public void Application_ViewActivated(object sender, ViewActivatedEventArgs e)
        //{
        //    // provide ExternalCommandData object to dockable page
        //    dockableWindow.CustomInitiator(edata);

        //}

        //// document opened event
        //private void Application_DocumentOpened(object sender, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs e)
        //{
        //    // provide ExternalCommandData object to dockable page
        //    dockableWindow.CustomInitiator(edata);
        //}

        public Result Execute(UIApplication app)
        {
            var data = new DockablePaneProviderData();
            var Viewer = new UI.Viewer();

            data.FrameworkElement = Viewer as FrameworkElement;

            // Setup initial state.
            var state = new DockablePaneState
            {
                DockPosition = DockPosition.Right,
            };

            var dpid = new DockablePaneId(new Guid("{ecea6d2f-533c-4e9d-a439-1c025aa0faee}"));

            app.RegisterDockablePane(dpid, "Crash Viewer", Viewer as IDockablePaneProvider);
            return Result.Succeeded;
        }

    }
}
