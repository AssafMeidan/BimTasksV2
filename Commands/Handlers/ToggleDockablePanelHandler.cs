using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Toggles the BimTasks dockable pane visibility.
    /// Uses DockablePane.Show() / Hide() via the pane ID registered during startup.
    /// </summary>
    public class ToggleDockablePanelHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var paneId = BimTasksApp.DockablePaneId;
                if (paneId == null)
                {
                    TaskDialog.Show("BimTasksV2", "Dockable pane not registered.");
                    return;
                }

                DockablePane pane = uiApp.GetDockablePane(paneId);
                if (pane == null)
                {
                    TaskDialog.Show("BimTasksV2", "Dockable pane not found.");
                    return;
                }

                if (pane.IsShown())
                    pane.Hide();
                else
                    pane.Show();

                Log.Information("ToggleDockablePanel: Pane is now {State}",
                    pane.IsShown() ? "visible" : "hidden");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ToggleDockablePanel failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
