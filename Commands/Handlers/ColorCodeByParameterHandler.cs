using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    public class ColorCodeByParameterHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                var eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                // Switch dockable panel to ColorCodeByParameter view
                eventAggregator.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                    .Publish("ColorCodeByParameter");

                // Signal the VM to scan for parameters
                eventAggregator.GetEvent<BimTasksEvents.ColorCodeInitEvent>().Publish(null);

                // Ensure the dockable pane is visible
                var pane = uiApp.GetDockablePane(BimTasksV2.Infrastructure.BimTasksBootstrapper.DockablePaneId);
                if (pane != null && !pane.IsShown())
                    pane.Show();

                Log.Information("ColorCodeByParameter: Published ColorCodeInitEvent");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ColorCodeByParameter handler failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
