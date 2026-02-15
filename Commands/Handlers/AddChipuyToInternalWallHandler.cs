using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Helpers.Cladding;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Adds cladding (chipuy) walls to the interior side only of selected walls.
    /// Validates the selection, then delegates to CladdingWallCreator.
    /// </summary>
    public class AddChipuyToInternalWallHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                List<Wall> selectedWalls = WallSelectionHelper.SelectWalls(uidoc);
                if (selectedWalls.Count == 0) return;

                CladdingWallCreator.AddCladdingToInternalSide(doc, selectedWalls);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "AddChipuyToInternalWallHandler failed");
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            }
        }
    }
}
