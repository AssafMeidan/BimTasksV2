using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers.DwgFloor;
using BimTasksV2.Views;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Interactive: user clicks on a DWG hatch face/edge, floor gets created from that boundary.
    /// </summary>
    public class PickAndCreateFloorHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Step 1: Pick a reference on a DWG element
                Reference pickedRef;
                try
                {
                    pickedRef = uidoc.Selection.PickObject(
                        ObjectType.PointOnElement,
                        new ImportInstanceFilter(),
                        "Pick a hatch or edge on a DWG import to create a floor boundary");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return;
                }

                if (pickedRef == null) return;

                // Step 2: Get the ImportInstance
                var element = doc.GetElement(pickedRef.ElementId);
                if (element is not ImportInstance importInstance)
                {
                    TaskDialog.Show("BimTasks", "Selected element is not a DWG import instance.");
                    return;
                }

                // Step 3: Try to get geometry from the reference
                List<Curve>? boundaryCurves = null;

                try
                {
                    var geomObj = element.GetGeometryObjectFromReference(pickedRef);
                    if (geomObj != null)
                    {
                        boundaryCurves = DwgGeometryExtractor.FindFaceAtReference(
                            importInstance, geomObj, pickedRef);
                    }
                }
                catch
                {
                    // GetGeometryObjectFromReference may fail for DWG sub-elements
                }

                // Fallback: find nearest face to picked point
                if (boundaryCurves == null && pickedRef.GlobalPoint != null)
                {
                    boundaryCurves = DwgGeometryExtractor.FindNearestFace(
                        importInstance, pickedRef.GlobalPoint);
                }

                if (boundaryCurves == null || boundaryCurves.Count < 3)
                {
                    TaskDialog.Show("BimTasks",
                        "Could not extract a valid floor boundary from the picked location.\n\n" +
                        "Try picking on a filled hatch region or a closed boundary edge in the DWG.");
                    return;
                }

                // Step 4: Build CurveLoop
                var curveLoop = FloorCreator.BuildCurveLoop(boundaryCurves);
                if (curveLoop == null)
                {
                    TaskDialog.Show("BimTasks",
                        "Could not build a valid closed curve loop from the extracted boundary.");
                    return;
                }

                // Step 5: Show floor type / level selector
                Level? defaultLevel = doc.ActiveView.GenLevel;
                var selector = new FloorTypeLevelSelector(doc, defaultLevel);
                if (selector.ShowDialog() != true ||
                    selector.SelectedFloorType == null ||
                    selector.SelectedLevel == null)
                    return;

                // Step 6: Create the floor
                using var trans = new Transaction(doc, "Create Floor from DWG");
                trans.Start();

                var floor = FloorCreator.CreateFloor(
                    doc, curveLoop, selector.SelectedFloorType.Id, selector.SelectedLevel.Id);

                if (floor != null)
                {
                    trans.Commit();
                    TaskDialog.Show("BimTasks", "Floor created successfully.");
                }
                else
                {
                    trans.RollBack();
                    TaskDialog.Show("BimTasks", "Failed to create floor from the boundary.");
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User pressed Esc
            }
            catch (Exception ex)
            {
                Log.Error(ex, "PickAndCreateFloorHandler failed");
                TaskDialog.Show("BimTasks Error", $"An error occurred:\n{ex.Message}");
            }
        }
    }
}
