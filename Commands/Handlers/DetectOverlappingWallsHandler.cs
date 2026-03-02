using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Helpers.OverlappingWallDetector;
using BimTasksV2.Views;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Detects walls of the same type that overlap each other (duplicate walls stacked at the same location).
    /// Shows results in a dialog where the user can highlight, unjoin, or delete the duplicates.
    /// </summary>
    public class DetectOverlappingWallsHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Use selected walls if any, otherwise all walls in active view
                var selected = WallSelectionHelper.GetSelectedWalls(uidoc);
                IList<Wall> walls;

                if (selected.Count > 0)
                {
                    walls = selected;
                    Log.Information("[DetectOverlappingWalls] Scanning {Count} selected walls", walls.Count);
                }
                else
                {
                    walls = new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .OfClass(typeof(Wall))
                        .WhereElementIsNotElementType()
                        .Cast<Wall>()
                        .ToList();
                    Log.Information("[DetectOverlappingWalls] Scanning {Count} walls in active view", walls.Count);
                }

                if (walls.Count < 2)
                {
                    TaskDialog.Show("BimTasksV2", "Need at least 2 walls to check for overlaps.\nSelect walls or open a view with walls.");
                    return;
                }

                var groups = OverlappingWallsDetector.FindOverlappingWalls(walls);

                if (groups.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", $"No overlapping same-type walls found among {walls.Count} walls.");
                    Log.Information("[DetectOverlappingWalls] No overlaps found");
                    return;
                }

                Log.Information("[DetectOverlappingWalls] Found {Count} overlapping groups ({Total} walls total)",
                    groups.Count, groups.Sum(g => g.Walls.Count));

                // Show results dialog
                var dialog = new OverlappingWallsResultDialog(groups, doc);
                dialog.ShowDialog();

                if (!dialog.Accepted)
                    return;

                switch (dialog.ChosenAction)
                {
                    case OverlappingWallAction.Highlight:
                        HighlightOverlapping(uidoc, groups);
                        break;

                    case OverlappingWallAction.Unjoin:
                        UnjoinOverlapping(doc, groups);
                        break;

                    case OverlappingWallAction.DeleteDuplicates:
                        DeleteDuplicates(doc, groups);
                        break;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "DetectOverlappingWalls failed");
                TaskDialog.Show("BimTasksV2", $"Error detecting overlapping walls: {ex.Message}");
            }
        }

        private static void HighlightOverlapping(UIDocument uidoc, List<OverlappingWallGroup> groups)
        {
            var allIds = groups.SelectMany(g => g.Walls.Select(w => w.Id)).ToList();
            uidoc.Selection.SetElementIds(allIds);

            // Temp isolate in view
            var view = uidoc.ActiveView;
            using var trans = new Transaction(uidoc.Document, "Highlight Overlapping Walls");
            trans.Start();
            try
            {
                view.IsolateElementsTemporary(allIds);
            }
            catch
            {
                // If isolation fails (e.g. sheet view), just select them
            }
            trans.Commit();

            Log.Information("[DetectOverlappingWalls] Highlighted {Count} walls in {Groups} groups",
                allIds.Count, groups.Count);
        }

        private static void UnjoinOverlapping(Document doc, List<OverlappingWallGroup> groups)
        {
            int unjoinCount = 0;

            using var trans = new Transaction(doc, "Unjoin Overlapping Walls");
            trans.Start();

            foreach (var group in groups)
            {
                for (int i = 0; i < group.Walls.Count; i++)
                {
                    for (int j = i + 1; j < group.Walls.Count; j++)
                    {
                        try
                        {
                            if (JoinGeometryUtils.AreElementsJoined(doc, group.Walls[i], group.Walls[j]))
                            {
                                JoinGeometryUtils.UnjoinGeometry(doc, group.Walls[i], group.Walls[j]);
                                unjoinCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Warning(ex, "[DetectOverlappingWalls] Failed to unjoin wall {A} and {B}",
                                group.Walls[i].Id, group.Walls[j].Id);
                        }
                    }
                }
            }

            trans.Commit();

            TaskDialog.Show("BimTasksV2", $"Unjoined {unjoinCount} wall pairs.\nThe duplicate walls should now be visible separately.");
            Log.Information("[DetectOverlappingWalls] Unjoined {Count} wall pairs", unjoinCount);
        }

        private static void DeleteDuplicates(Document doc, List<OverlappingWallGroup> groups)
        {
            int deleteCount = 0;
            var toDelete = new List<ElementId>();

            // For each group, keep the first wall (oldest by element ID) and delete the rest
            foreach (var group in groups)
            {
                var sorted = group.Walls.OrderBy(w => w.Id.Value).ToList();
                for (int i = 1; i < sorted.Count; i++)
                {
                    toDelete.Add(sorted[i].Id);
                }
            }

            using var trans = new Transaction(doc, "Delete Duplicate Walls");
            trans.Start();

            foreach (var id in toDelete)
            {
                try
                {
                    doc.Delete(id);
                    deleteCount++;
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[DetectOverlappingWalls] Failed to delete wall {Id}", id);
                }
            }

            trans.Commit();

            TaskDialog.Show("BimTasksV2", $"Deleted {deleteCount} duplicate walls from {groups.Count} overlapping groups.");
            Log.Information("[DetectOverlappingWalls] Deleted {Count} duplicate walls", deleteCount);
        }

        /// <summary>
        /// Temporary diagnostic: fetches two walls by ID and dumps every property
        /// the detector checks, step by step, to find why detection fails.
        /// </summary>
        private static void DiagnoseWallPair(Document doc, long id1, long id2)
        {
            var el1 = doc.GetElement(new ElementId(id1));
            var el2 = doc.GetElement(new ElementId(id2));

            var lines = new List<string> { $"=== DIAGNOSTIC: Wall {id1} vs {id2} ===" };

            if (el1 == null) { lines.Add($"Wall {id1}: NOT FOUND in document"); }
            if (el2 == null) { lines.Add($"Wall {id2}: NOT FOUND in document"); }
            if (el1 == null || el2 == null)
            {
                TaskDialog.Show("BimTasksV2 - Diagnostic", string.Join("\n", lines));
                return;
            }

            var w1 = el1 as Wall;
            var w2 = el2 as Wall;
            if (w1 == null) lines.Add($"{id1} is NOT a Wall — it's {el1.GetType().Name} / Category: {el1.Category?.Name}");
            if (w2 == null) lines.Add($"{id2} is NOT a Wall — it's {el2.GetType().Name} / Category: {el2.Category?.Name}");
            if (w1 == null || w2 == null)
            {
                TaskDialog.Show("BimTasksV2 - Diagnostic", string.Join("\n", lines));
                return;
            }

            // 1. Wall Type
            lines.Add($"\n--- WALL TYPE ---");
            lines.Add($"W1 type: {w1.WallType.Name} (TypeId={w1.WallType.Id.Value})");
            lines.Add($"W2 type: {w2.WallType.Name} (TypeId={w2.WallType.Id.Value})");
            lines.Add($"Same type? {(w1.WallType.Id.Value == w2.WallType.Id.Value ? "YES" : "NO")}");

            // 2. Location
            lines.Add($"\n--- LOCATION ---");
            var loc1 = w1.Location as LocationCurve;
            var loc2 = w2.Location as LocationCurve;
            lines.Add($"W1 has LocationCurve? {loc1 != null}");
            lines.Add($"W2 has LocationCurve? {loc2 != null}");

            if (loc1 != null && loc2 != null)
            {
                var curve1 = loc1.Curve;
                var curve2 = loc2.Curve;
                lines.Add($"W1 curve type: {curve1.GetType().Name}");
                lines.Add($"W2 curve type: {curve2.GetType().Name}");

                var p1a = curve1.GetEndPoint(0);
                var p1b = curve1.GetEndPoint(1);
                var p2a = curve2.GetEndPoint(0);
                var p2b = curve2.GetEndPoint(1);

                lines.Add($"W1 start: ({p1a.X:F6}, {p1a.Y:F6}, {p1a.Z:F6})");
                lines.Add($"W1 end:   ({p1b.X:F6}, {p1b.Y:F6}, {p1b.Z:F6})");
                lines.Add($"W1 length: {curve1.Length:F6} ft ({curve1.Length * 304.8:F1} mm)");
                lines.Add($"W2 start: ({p2a.X:F6}, {p2a.Y:F6}, {p2a.Z:F6})");
                lines.Add($"W2 end:   ({p2b.X:F6}, {p2b.Y:F6}, {p2b.Z:F6})");
                lines.Add($"W2 length: {curve2.Length:F6} ft ({curve2.Length * 304.8:F1} mm)");

                // 3. Parallel check (only for lines)
                if (curve1 is Line line1 && curve2 is Line line2)
                {
                    var dir1 = line1.Direction;
                    var dir2 = line2.Direction;
                    double cross = Math.Abs(dir1.X * dir2.Y - dir1.Y * dir2.X);
                    lines.Add($"\n--- PARALLEL CHECK ---");
                    lines.Add($"W1 dir: ({dir1.X:F6}, {dir1.Y:F6}, {dir1.Z:F6})");
                    lines.Add($"W2 dir: ({dir2.X:F6}, {dir2.Y:F6}, {dir2.Z:F6})");
                    lines.Add($"Cross product (XY): {cross:F8}  (threshold: 0.018)");
                    lines.Add($"Parallel? {(cross <= 0.018 ? "YES" : "NO")}");

                    // 4. Perpendicular distance
                    var delta = p2a - p1a;
                    var perp = delta - dir1.DotProduct(delta) * dir1;
                    double perpDist = Math.Sqrt(perp.X * perp.X + perp.Y * perp.Y);
                    lines.Add($"\n--- DISTANCE CHECK ---");
                    lines.Add($"Delta (p2a - p1a): ({delta.X:F6}, {delta.Y:F6}, {delta.Z:F6})");
                    lines.Add($"Perp component: ({perp.X:F6}, {perp.Y:F6}, {perp.Z:F6})");
                    lines.Add($"Perp distance: {perpDist:F8} ft ({perpDist * 304.8:F3} mm)  (threshold: 0.016 ft / 5 mm)");
                    lines.Add($"Close enough? {(perpDist <= 0.016 ? "YES" : "NO")}");

                    // 5. Level check
                    var level1 = w1.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
                    var level2 = w2.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
                    lines.Add($"\n--- LEVEL CHECK ---");
                    lines.Add($"W1 base level: {level1?.Value}");
                    lines.Add($"W2 base level: {level2?.Value}");
                    lines.Add($"Same level? {(level1 != null && level2 != null ? (level1 == level2 ? "YES" : "NO") : "N/A")}");

                    // 6. Overlap span
                    double a1 = dir1.DotProduct(p1a);
                    double b1 = dir1.DotProduct(p1b);
                    double a2 = dir1.DotProduct(p2a);
                    double b2 = dir1.DotProduct(p2b);
                    double min1 = Math.Min(a1, b1), max1 = Math.Max(a1, b1);
                    double min2 = Math.Min(a2, b2), max2 = Math.Max(a2, b2);
                    double overlapStart = Math.Max(min1, min2);
                    double overlapEnd = Math.Min(max1, max2);
                    double overlapLen = Math.Max(0, overlapEnd - overlapStart);
                    lines.Add($"\n--- OVERLAP SPAN ---");
                    lines.Add($"W1 projected: [{min1:F6}, {max1:F6}]");
                    lines.Add($"W2 projected: [{min2:F6}, {max2:F6}]");
                    lines.Add($"Overlap: [{overlapStart:F6}, {overlapEnd:F6}] = {overlapLen:F6} ft ({overlapLen * 304.8:F1} mm)");
                    lines.Add($"Overlap >= 50mm? {(overlapLen >= 0.164 ? "YES" : "NO")}");
                }
                else
                {
                    lines.Add($"\nSKIPPED: One or both curves are not Line (detector only handles linear walls)");
                }
            }

            // 7. Join status
            lines.Add($"\n--- JOIN STATUS ---");
            try
            {
                bool joined = JoinGeometryUtils.AreElementsJoined(doc, w1, w2);
                lines.Add($"Geometry joined? {(joined ? "YES" : "NO")}");
            }
            catch (Exception ex)
            {
                lines.Add($"Join check failed: {ex.Message}");
            }

            // 8. Width
            lines.Add($"\n--- WIDTH ---");
            lines.Add($"W1 width: {w1.Width:F6} ft ({w1.Width * 304.8:F1} mm)");
            lines.Add($"W2 width: {w2.Width:F6} ft ({w2.Width * 304.8:F1} mm)");

            var msg = string.Join("\n", lines);
            Log.Information("[DiagnoseWallPair]\n{Msg}", msg);
            TaskDialog.Show("BimTasksV2 - Wall Pair Diagnostic", msg);
        }
    }
}
