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

            // For each group, keep the wall with the most hosted elements (doors, windows, etc.)
            // and delete the rest
            foreach (var group in groups)
            {
                var best = group.Walls
                    .OrderByDescending(w => CountHostedElements(doc, w))
                    .ThenBy(w => w.Id.Value) // tie-break: keep oldest
                    .First();

                foreach (var w in group.Walls)
                {
                    if (w.Id != best.Id)
                        toDelete.Add(w.Id);
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
        /// Counts hosted elements (doors, windows, openings, inserts) on a wall.
        /// </summary>
        private static int CountHostedElements(Document doc, Wall wall)
        {
            var filter = new ElementFilter[]
            {
                new ElementCategoryFilter(BuiltInCategory.OST_Doors),
                new ElementCategoryFilter(BuiltInCategory.OST_Windows),
                new ElementCategoryFilter(BuiltInCategory.OST_GenericModel),
            };
            var orFilter = new LogicalOrFilter(filter);

            return new FilteredElementCollector(doc)
                .WherePasses(orFilter)
                .WhereElementIsNotElementType()
                .Where(e => e is FamilyInstance fi && fi.Host?.Id == wall.Id)
                .Count();
        }
    }
}
