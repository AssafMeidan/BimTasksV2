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
    /// Joins geometry between all concrete walls and selected floor/slab/foundation/roof/ceiling elements.
    /// Uses SuppressWarningsPreprocessor to handle Revit join warnings silently.
    /// Supports floors, roofs, ceilings, and structural foundations.
    /// Excludes curtain walls, stacked walls, and walls thinner than 30mm.
    /// Tries alternative join order when first attempt fails.
    /// </summary>
    public class JoinAllConcreteWallsAndFloorsGeometryHandler : ICommandHandler
    {
        private const string CommandName = nameof(JoinAllConcreteWallsAndFloorsGeometryHandler);
        private const double MinWallWidthFeet = 0.1; // ~30mm

        private static readonly BuiltInCategory[] SupportedHostCategories =
        {
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Ceilings
        };

        private class JoinStats
        {
            public int HostsProcessed { get; set; }
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
                        "Please select floors, roofs, ceilings, or foundations.");
                    return;
                }

                // Get host elements
                var hostElements = GetHostElements(doc, selectedIds);
                if (hostElements.Count == 0)
                {
                    TaskDialog.Show("No Valid Elements",
                        "No floors, roofs, ceilings, or foundations found in selection.");
                    return;
                }

                Log.Information("[{Command}] Found {Count} host elements", CommandName, hostElements.Count);

                // Execute join operation
                using (var trans = new Transaction(doc, "Join Walls to Floors"))
                {
                    ConfigureFailureHandling(trans);
                    trans.Start();
                    ProcessAllHosts(doc, hostElements, stats);
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

        private List<Element> GetHostElements(Document doc, ICollection<ElementId> selectedIds)
        {
            var filter = new ElementMulticategoryFilter(SupportedHostCategories);
            return new FilteredElementCollector(doc, selectedIds)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToList();
        }

        private void ProcessAllHosts(Document doc, List<Element> hostElements, JoinStats stats)
        {
            foreach (var host in hostElements)
            {
                if (!IsValidElement(host)) continue;
                stats.HostsProcessed++;
                ProcessSingleHost(doc, host, stats);
            }
        }

        private void ProcessSingleHost(Document doc, Element host, JoinStats stats)
        {
            List<Wall> walls;
            try
            {
                var intersectFilter = new ElementIntersectsElementFilter(host);
                walls = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall))
                    .WherePasses(intersectFilter)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
            }
            catch (Exception ex)
            {
                Log.Warning("[{Command}] Failed to get intersecting walls for host {Id}: {Error}",
                    CommandName, host.Id.Value, ex.Message);
                return;
            }

            foreach (var wall in walls)
            {
                ProcessWall(doc, host, wall, stats);
            }
        }

        private void ProcessWall(Document doc, Element host, Wall wall, JoinStats stats)
        {
            if (!IsValidElement(wall))
            {
                stats.SkippedInvalid++;
                return;
            }

            try
            {
                if (JoinGeometryUtils.AreElementsJoined(doc, host, wall))
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

            TryJoinWall(doc, host, wall, stats);
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

        private void TryJoinWall(Document doc, Element host, Wall wall, JoinStats stats)
        {
            // Attempt 1: Normal order
            try
            {
                JoinGeometryUtils.JoinGeometry(doc, host, wall);
                stats.WallsJoined++;
                return;
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Try switched order
            }
            catch (Exception ex)
            {
                RecordFailure(stats, host, wall, ex.Message);
                return;
            }

            // Attempt 2: Switched order
            try
            {
                JoinGeometryUtils.JoinGeometry(doc, wall, host);
                stats.WallsJoinedSwitched++;
            }
            catch (Exception ex)
            {
                RecordFailure(stats, host, wall, ex.Message);
            }
        }

        private void RecordFailure(JoinStats stats, Element host, Wall wall, string reason)
        {
            stats.Failed++;
            if (stats.FailureMessages.Count < 30)
            {
                stats.FailureMessages.Add($"Wall {wall.Id.Value} <-> Host {host.Id.Value}");
            }
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
                MainInstruction = $"Processed {stats.HostsProcessed} elements in {elapsedMs}ms",
                MainContent =
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
