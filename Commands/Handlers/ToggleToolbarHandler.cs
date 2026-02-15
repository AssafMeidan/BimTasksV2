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
    /// Publishes ToggleFloatingToolbarEvent to show/hide the floating toolbar window.
    /// </summary>
    public class ToggleToolbarHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                eventAggregator.GetEvent<BimTasksEvents.ToggleFloatingToolbarEvent>().Publish(null);

                Log.Information("ToggleToolbar: Published ToggleFloatingToolbarEvent");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ToggleToolbar failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
