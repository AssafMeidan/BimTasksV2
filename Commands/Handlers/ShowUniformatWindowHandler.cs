using System;
using System.Windows;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Creates and shows a modeless UniformatWindowView.
    /// The view is resolved from the container so ViewModelLocator auto-wires the VM.
    /// </summary>
    public class ShowUniformatWindowHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;

                // Update context service so the VM can access the Revit document
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                // Resolve the view (ViewModelLocator will auto-wire the VM)
                var win = container.Resolve<Views.UniformatWindowView>();

                // Set owner to Revit main window if possible
                try
                {
                    if (Application.Current?.MainWindow != null)
                        win.Owner = Application.Current.MainWindow;
                }
                catch { /* Owner assignment may fail in some contexts */ }

                // Show modeless
                win.Show();

                Log.Information("ShowUniformatWindow: Window shown (modeless)");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ShowUniformatWindow failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
