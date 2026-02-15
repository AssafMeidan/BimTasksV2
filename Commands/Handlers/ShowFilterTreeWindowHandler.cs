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
    /// Publishes ResetFilterTreeEvent and switches the dockable panel to FilterTreeView.
    /// The FilterTreeViewModel subscribes to the event and rebuilds the tree.
    /// </summary>
    public class ShowFilterTreeWindowHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                // Update context service
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                var eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                // Switch dockable panel to FilterTreeView
                eventAggregator.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                    .Publish("FilterTree");

                // Reset the filter tree
                eventAggregator.GetEvent<BimTasksEvents.ResetFilterTreeEvent>().Publish(null);

                // Ensure the dockable pane is visible
                var paneId = BimTasksApp.DockablePaneId;
                if (paneId != null)
                {
                    DockablePane pane = uiApp.GetDockablePane(paneId);
                    if (pane != null && !pane.IsShown())
                        pane.Show();
                }

                Log.Information("ShowFilterTreeWindow: Published ResetFilterTreeEvent");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ShowFilterTreeWindow failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
