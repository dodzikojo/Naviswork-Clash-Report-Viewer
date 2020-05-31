using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace CrashViewerRevitAddIn
{
    
    
    public class MainClass : IExternalApplication
    {
       
        public Result OnStartup(UIControlledApplication application)
        {
            //Ribbon Panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(Tab.AddIns, "Crash Viewer");

            //Assembly
            Assembly assembly = Assembly.GetExecutingAssembly();

            //Assembly Path
            string assemblyPath = assembly.Location;

            // Create Register Button
            //PushButton registerButton = ribbonPanel.AddItem(new PushButtonData("Register Window", "Register",
            //    assemblyPath, "CrashViewerRevitAddIn.Main.Register")) as PushButton;
            //// accessibility check for register Button
           // registerButton.AvailabilityClassName = "CrashViewerRevitAddIn.CommandAvailability";
            // btn tooltip 
            //registerButton.ToolTip = "Register dockable window at the zero document state.";
            // register button icon images
            //registerButton.LargeImage = GetResourceImage(assembly, "CrashViewerRevitAddIn.Resources.register32.png");
            //registerButton.Image = GetResourceImage(assembly, "CrashViewerRevitAddIn.Resources.register16.png");

            // Create Show Button
            PushButton showButton = ribbonPanel.AddItem(new PushButtonData("Show Window", "Show Crash Viewer", assemblyPath,
                "CrashViewerRevitAddIn.Main.Show")) as PushButton;

            //registerButton.AvailabilityClassName = "CrashViewerRevitAddIn.CommandAvailability";
            // btn tooltip
            showButton.ToolTip = "Show Crash Viewer if it's not currently visible.";
            // show button icon images
            showButton.LargeImage = GetResourceImage(assembly, "CrashViewerRevitAddIn.Resources.show32.png");
            showButton.Image = GetResourceImage(assembly, "CrashViewerRevitAddIn.Resources.show16.png");

            application.ControlledApplication.ApplicationInitialized += DockablePaneRegisters;

            return Result.Succeeded;
        }


        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private void DockablePaneRegisters(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            // Register dockable pane.
            var registercommand = new Main.Register();
            registercommand.Execute(new UIApplication(sender as Application));
        }


        public ImageSource GetResourceImage(Assembly assembly, string imageName)
        {
            try
            {
                //bitmap stream to construct bitmap frame
                Stream resource = assembly.GetManifestResourceStream(imageName);
                //return image data
                return BitmapFrame.Create(resource);
            }
            catch (Exception)
            {

                return null;
            }
        }
    }

    // external command availability
    public class CommandAvailability : IExternalCommandAvailability
    {
        // interface member method
        public bool IsCommandAvailable(UIApplication app, CategorySet cate)
        {
            // zero doc state
            if (app.ActiveUIDocument == null)
            {
                // disable register btn
                return true;
            }
            // enable register btn
            return true;
        }
    }


}
