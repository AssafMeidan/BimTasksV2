using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    public class ColorSwatchHandler : ICommandHandler
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

                // Switch dockable panel to ColorSwatchView
                eventAggregator.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                    .Publish("ColorSwatch");

                // Initialize the ViewModel with current Revit context
                eventAggregator.GetEvent<BimTasksEvents.InitializeColorSwatchEvent>()
                    .Publish(uiApp);

                // Ensure the dockable pane is visible
                var pane = uiApp.GetDockablePane(BimTasksV2.Infrastructure.BimTasksBootstrapper.DockablePaneId);
                if (pane != null && !pane.IsShown())
                    pane.Show();

                Log.Information("[ColorSwatchHandler] Opened color swatch panel");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[ColorSwatchHandler] Failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
