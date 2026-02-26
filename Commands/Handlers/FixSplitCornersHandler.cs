using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers.WallSplitter;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Phase 2 of wall splitting: reads saved corner data and trims/extends
    /// replacement walls to form clean L/T corner intersections.
    /// Run this after Split Wall and after dismissing any join errors.
    /// </summary>
    public class FixSplitCornersHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var doc = uiApp.ActiveUIDocument.Document;

            try
            {
                string result = WallSplitterEngine.FixCorners(doc);
                TaskDialog.Show("BimTasksV2 - Fix Split Corners", result);
                Log.Information("[FixSplitCornersHandler] {Result}", result);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "FixSplitCorners failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
