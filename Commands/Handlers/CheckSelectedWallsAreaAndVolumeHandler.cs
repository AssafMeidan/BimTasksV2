using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Publishes CalculateElementsEvent and switches the dockable panel to ElementCalculationView.
    /// The ElementCalculationViewModel subscribes to the event and performs the actual calculation.
    /// </summary>
    public class CheckSelectedWallsAreaAndVolumeHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                // Update context service so the VM can access the active selection
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                var eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                // Switch dockable panel to ElementCalculationView
                eventAggregator.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                    .Publish("ElementCalculation");

                // Trigger calculation
                eventAggregator.GetEvent<BimTasksEvents.CalculateElementsEvent>().Publish(null);

                // Ensure the dockable pane is visible
                var paneId = BimTasksApp.DockablePaneId;
                if (paneId != null)
                {
                    DockablePane pane = uiApp.GetDockablePane(paneId);
                    if (pane != null && !pane.IsShown())
                        pane.Show();
                }

                Log.Information("CheckSelectedWallsAreaAndVolume: Published CalculateElementsEvent");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CheckSelectedWallsAreaAndVolume failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
