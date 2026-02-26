using System;
using System.Collections.Generic;
using System.Linq;
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
    /// Batch: user picks DWG element, selects floor type and level, ALL floor-sized hatches get imported.
    /// </summary>
    public class ImportAllDwgFloorsHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Step 1: Pick a DWG ImportInstance
                Reference pickedRef;
                try
                {
                    pickedRef = uidoc.Selection.PickObject(
                        ObjectType.Element,
                        new ImportInstanceFilter(),
                        "Select a DWG import/link to import all floor boundaries from");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return;
                }

                if (pickedRef == null) return;

                var element = doc.GetElement(pickedRef.ElementId);
                if (element is not ImportInstance importInstance)
                {
                    TaskDialog.Show("BimTasks", "Selected element is not a DWG import instance.");
                    return;
                }

                // Step 2: Show floor type / level selector
                Level? defaultLevel = doc.ActiveView.GenLevel;
                var selector = new FloorTypeLevelSelector(doc, defaultLevel);
                if (selector.ShowDialog() != true ||
                    selector.SelectedFloorType == null ||
                    selector.SelectedLevel == null)
                    return;

                var floorTypeId = selector.SelectedFloorType.Id;
                var levelId = selector.SelectedLevel.Id;

                // Step 3: Extract visible geometry from the DWG (filtered by active view)
                var extraction = DwgGeometryExtractor.ExtractAll(importInstance, doc.ActiveView);

                // Step 4: Build CurveLoops from face loops
                var validLoops = new List<CurveLoop>();
                int skippedArea = 0;
                int skippedInvalid = 0;

                foreach (var faceLoop in extraction.FaceLoops)
                {
                    var curveLoop = FloorCreator.BuildCurveLoop(faceLoop);
                    if (curveLoop == null)
                    {
                        skippedInvalid++;
                        continue;
                    }

                    double areaSqft = FloorCreator.ComputeAreaSqft(curveLoop);
                    if (areaSqft < DwgFloorConfig.MinAreaSqft)
                    {
                        skippedArea++;
                        continue;
                    }

                    validLoops.Add(curveLoop);
                }

                // Step 5: Chain loose curves into additional loops
                if (extraction.LooseCurves.Count > 0)
                {
                    var chainedLoops = CurveChainer.ChainCurves(extraction.LooseCurves);
                    foreach (var chainedCurves in chainedLoops)
                    {
                        var curveLoop = FloorCreator.BuildCurveLoop(chainedCurves);
                        if (curveLoop == null)
                        {
                            skippedInvalid++;
                            continue;
                        }

                        double areaSqft = FloorCreator.ComputeAreaSqft(curveLoop);
                        if (areaSqft < DwgFloorConfig.MinAreaSqft)
                        {
                            skippedArea++;
                            continue;
                        }

                        validLoops.Add(curveLoop);
                    }
                }

                // Step 6: Deduplicate loops with matching centroid + area
                int skippedDuplicate = 0;
                var uniqueLoops = new List<CurveLoop>();
                var loopSignatures = new List<(XYZ Centroid, double Area)>();

                foreach (var loop in validLoops)
                {
                    double area = FloorCreator.ComputeAreaSqft(loop);
                    var centroid = FloorCreator.ComputeCentroid(loop);

                    bool isDuplicate = false;
                    foreach (var (existingCentroid, existingArea) in loopSignatures)
                    {
                        if (Math.Abs(area - existingArea) < 1.0 &&
                            centroid.DistanceTo(existingCentroid) < DwgFloorConfig.ClosureTolerance)
                        {
                            isDuplicate = true;
                            break;
                        }
                    }

                    if (isDuplicate)
                    {
                        skippedDuplicate++;
                        continue;
                    }

                    uniqueLoops.Add(loop);
                    loopSignatures.Add((centroid, area));
                }

                validLoops = uniqueLoops;

                if (validLoops.Count == 0)
                {
                    TaskDialog.Show("BimTasks - No Floors Found",
                        $"No valid floor boundaries found in the DWG.\n\n" +
                        $"Face loops found: {extraction.FaceLoops.Count}\n" +
                        $"Loose curves found: {extraction.LooseCurves.Count}\n" +
                        $"Skipped (too small < {DwgFloorConfig.MinAreaSqm} m\u00b2): {skippedArea}\n" +
                        $"Skipped (duplicates): {skippedDuplicate}\n" +
                        $"Skipped (invalid geometry): {skippedInvalid}\n" +
                        $"Geometry errors: {extraction.ErrorCount}");
                    return;
                }

                // Step 7: Confirm with user
                var confirmResult = TaskDialog.Show("BimTasks - Confirm Import",
                    $"Found {validLoops.Count} unique floor boundary(ies).\n\n" +
                    $"Skipped (too small < {DwgFloorConfig.MinAreaSqm} m\u00b2): {skippedArea}\n" +
                    $"Skipped (duplicates): {skippedDuplicate}\n" +
                    $"Skipped (invalid geometry): {skippedInvalid}\n" +
                    $"Geometry errors: {extraction.ErrorCount}\n\n" +
                    "Proceed with floor creation?",
                    TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No);

                if (confirmResult != TaskDialogResult.Yes)
                    return;

                // Step 7: Create all floors
                var result = FloorCreator.CreateFloors(doc, validLoops, floorTypeId, levelId);

                // Step 8: Show results
                string resultMsg = $"Floors created: {result.Created}\nFloors failed: {result.Failed}";

                if (result.Errors.Count > 0)
                {
                    resultMsg += "\n\nErrors:\n" +
                                 string.Join("\n", result.Errors.Take(10));
                    if (result.Errors.Count > 10)
                        resultMsg += $"\n... and {result.Errors.Count - 10} more";
                }

                TaskDialog.Show("BimTasks - Import Results", resultMsg);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User pressed Esc
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ImportAllDwgFloorsHandler failed");
                TaskDialog.Show("BimTasks Error", $"An error occurred:\n{ex.Message}");
            }
        }
    }
}
