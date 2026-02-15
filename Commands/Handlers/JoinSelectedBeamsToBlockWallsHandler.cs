using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Joins selected structural elements (beams, columns, foundations) to intersecting walls.
    /// Uses SuppressWarningsPreprocessor to handle Revit join warnings silently.
    /// Excludes curtain walls, stacked walls, and walls thinner than 30mm.
    /// Tries alternative join order when first attempt fails.
    /// </summary>
    public class JoinSelectedBeamsToBlockWallsHandler : ICommandHandler
    {
        private const string CommandName = nameof(JoinSelectedBeamsToBlockWallsHandler);
        private const double MinWallWidthFeet = 0.1; // ~30mm

        private static readonly BuiltInCategory[] SupportedStructuralCategories =
        {
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFoundation
        };

        private class JoinStats
        {
            public int StructuralElementsProcessed { get; set; }
            public int BeamsProcessed { get; set; }
            public int ColumnsProcessed { get; set; }
            public int FoundationsProcessed { get; set; }
            public int WallsJoined { get; set; }
            public int WallsJoinedSwitched { get; set; }
            public int AlreadyJoined { get; set; }
            public int SkippedTooThin { get; set; }
            public int SkippedCurtain { get; set; }
            public int SkippedInvalid { get; set; }
            public int Failed { get; set; }
            public List<string> FailureMessages { get; } = new List<string>();
            public int TotalJoined => WallsJoined + WallsJoinedSwitched;
        }

        public void Execute(UIApplication uiApp)
        {
            var stopwatch = Stopwatch.StartNew();
            Log.Information("[{Command}] Starting execution", CommandName);

            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var stats = new JoinStats();

            try
            {
                // Validate selection
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("No Selection",
                        "Please select beams, columns, or foundations.");
                    return;
                }

                // Get structural elements
                var structuralElements = GetStructuralElements(doc, selectedIds, stats);
                if (structuralElements.Count == 0)
                {
                    TaskDialog.Show("No Valid Elements",
                        "No beams, columns, or foundations found in selection.");
                    return;
                }

                Log.Information("[{Command}] Found {Count} structural elements", CommandName, structuralElements.Count);

                // Reset counts (they were used for initial logging only)
                stats.BeamsProcessed = 0;
                stats.ColumnsProcessed = 0;
                stats.FoundationsProcessed = 0;

                // Execute join operation
                using (var trans = new Transaction(doc, "Join Walls to Beams/Columns"))
                {
                    ConfigureFailureHandling(trans);
                    trans.Start();
                    ProcessAllStructuralElements(doc, structuralElements, stats);
                    trans.Commit();
                }

                stopwatch.Stop();
                ShowResultsDialog(stats, stopwatch.ElapsedMilliseconds);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                Log.Information("[{Command}] Cancelled by user", CommandName);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Log.Error(ex, "[{Command}] Fatal error after {Elapsed}ms", CommandName, stopwatch.ElapsedMilliseconds);
                TaskDialog.Show("Error", $"An error occurred: {ex.Message}");
            }
        }

        private List<Element> GetStructuralElements(Document doc, ICollection<ElementId> selectedIds, JoinStats stats)
        {
            var filter = new ElementMulticategoryFilter(SupportedStructuralCategories);
            var elements = new FilteredElementCollector(doc, selectedIds)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToList();

            foreach (var element in elements)
            {
                var catId = element.Category?.Id.Value ?? 0;
                if (catId == (long)BuiltInCategory.OST_StructuralFraming)
                    stats.BeamsProcessed++;
                else if (catId == (long)BuiltInCategory.OST_StructuralColumns)
                    stats.ColumnsProcessed++;
                else if (catId == (long)BuiltInCategory.OST_StructuralFoundation)
                    stats.FoundationsProcessed++;
            }

            return elements;
        }

        private void ProcessAllStructuralElements(Document doc, List<Element> structuralElements, JoinStats stats)
        {
            foreach (var structElement in structuralElements)
            {
                if (!IsValidElement(structElement))
                {
                    stats.SkippedInvalid++;
                    continue;
                }

                stats.StructuralElementsProcessed++;
                TrackCategory(structElement, stats);
                ProcessSingleStructuralElement(doc, structElement, stats);
            }
        }

        private void TrackCategory(Element element, JoinStats stats)
        {
            var catId = element.Category?.Id.Value ?? 0;
            if (catId == (long)BuiltInCategory.OST_StructuralFraming)
                stats.BeamsProcessed++;
            else if (catId == (long)BuiltInCategory.OST_StructuralColumns)
                stats.ColumnsProcessed++;
            else if (catId == (long)BuiltInCategory.OST_StructuralFoundation)
                stats.FoundationsProcessed++;
        }

        private void ProcessSingleStructuralElement(Document doc, Element structElement, JoinStats stats)
        {
            List<Wall> walls;
            try
            {
                var intersectFilter = new ElementIntersectsElementFilter(structElement);
                walls = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall))
                    .WherePasses(intersectFilter)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
            }
            catch (Exception ex)
            {
                Log.Warning("[{Command}] Failed to get intersecting walls for {Id}: {Error}",
                    CommandName, structElement.Id.Value, ex.Message);
                return;
            }

            foreach (var wall in walls)
            {
                ProcessWall(doc, structElement, wall, stats);
            }
        }

        private void ProcessWall(Document doc, Element structElement, Wall wall, JoinStats stats)
        {
            if (!IsValidElement(wall))
            {
                stats.SkippedInvalid++;
                return;
            }

            try
            {
                if (JoinGeometryUtils.AreElementsJoined(doc, structElement, wall))
                {
                    stats.AlreadyJoined++;
                    return;
                }
            }
            catch
            {
                stats.SkippedInvalid++;
                return;
            }

            if (wall.Width < MinWallWidthFeet)
            {
                stats.SkippedTooThin++;
                return;
            }

            if (IsCurtainOrStackedWall(wall))
            {
                stats.SkippedCurtain++;
                return;
            }

            TryJoinWall(doc, structElement, wall, stats);
        }

        private bool IsCurtainOrStackedWall(Wall wall)
        {
            try
            {
                var wallType = wall.WallType;
                if (wallType == null) return false;

                if (wallType.Kind == WallKind.Curtain || wallType.Kind == WallKind.Stacked)
                    return true;

                string familyName = wallType.FamilyName ?? "";
                if (familyName.IndexOf("Curtain", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    familyName.IndexOf("Storefront", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        private void TryJoinWall(Document doc, Element structElement, Wall wall, JoinStats stats)
        {
            try
            {
                JoinGeometryUtils.JoinGeometry(doc, structElement, wall);
                stats.WallsJoined++;
                return;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Try switched order
            }
            catch (Exception ex)
            {
                RecordFailure(stats, structElement, wall, ex.Message);
                return;
            }

            try
            {
                JoinGeometryUtils.JoinGeometry(doc, wall, structElement);
                stats.WallsJoinedSwitched++;
            }
            catch (Exception ex)
            {
                RecordFailure(stats, structElement, wall, ex.Message);
            }
        }

        private void RecordFailure(JoinStats stats, Element structElement, Wall wall, string reason)
        {
            stats.Failed++;
            string categoryName = GetCategoryShortName(structElement);
            if (stats.FailureMessages.Count < 30)
            {
                stats.FailureMessages.Add($"Wall {wall.Id.Value} <-> {categoryName} {structElement.Id.Value}");
            }
        }

        private static string GetCategoryShortName(Element element)
        {
            var catId = element.Category?.Id.Value ?? 0;
            if (catId == (long)BuiltInCategory.OST_StructuralFraming) return "Beam";
            if (catId == (long)BuiltInCategory.OST_StructuralColumns) return "Column";
            if (catId == (long)BuiltInCategory.OST_StructuralFoundation) return "Foundation";
            return element.Category?.Name ?? "Unknown";
        }

        private static bool IsValidElement(Element element)
        {
            try
            {
                return element != null &&
                       element.IsValidObject &&
                       element.Id != ElementId.InvalidElementId;
            }
            catch
            {
                return false;
            }
        }

        private static void ConfigureFailureHandling(Transaction trans)
        {
            var options = trans.GetFailureHandlingOptions();
            options.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
            options.SetClearAfterRollback(true);
            trans.SetFailureHandlingOptions(options);
        }

        private void ShowResultsDialog(JoinStats stats, long elapsedMs)
        {
            var dialog = new TaskDialog("Join Operation Complete")
            {
                MainInstruction = $"Processed {stats.StructuralElementsProcessed} structural elements in {elapsedMs}ms",
                MainContent =
                    $"Structural Elements:\n" +
                    $"   Beams: {stats.BeamsProcessed}\n" +
                    $"   Columns: {stats.ColumnsProcessed}\n" +
                    $"   Foundations: {stats.FoundationsProcessed}\n\n" +
                    $"Walls joined: {stats.TotalJoined}\n" +
                    $"   (Normal: {stats.WallsJoined}, Switched: {stats.WallsJoinedSwitched})\n\n" +
                    $"Already joined: {stats.AlreadyJoined}\n\n" +
                    $"Skipped (thin): {stats.SkippedTooThin}\n\n" +
                    $"Skipped (curtain): {stats.SkippedCurtain}\n\n" +
                    $"Failed: {stats.Failed}"
            };

            if (stats.FailureMessages.Any())
            {
                dialog.ExpandedContent = "Failed joins:\n" + string.Join("\n", stats.FailureMessages);
                if (stats.Failed > stats.FailureMessages.Count)
                {
                    dialog.ExpandedContent += $"\n... and {stats.Failed - stats.FailureMessages.Count} more";
                }
            }

            dialog.Show();
        }
    }
}
