using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Helpers.WallSplitter;
using BimTasksV2.Views;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Splits compound walls into individual single-layer walls.
    /// Saves split results to a JSON file and shows the dockable panel
    /// with a Trim Corners view. The user OKs Revit warnings, then clicks
    /// Trim Corners as a separate command.
    /// </summary>
    public class SplitWallHandler : ICommandHandler
    {
        /// <summary>
        /// Path to the JSON file storing split results for the Trim Corners command.
        /// Stored next to the plugin DLL so both commands can find it.
        /// </summary>
        public static string SplitResultsPath =>
            Path.Combine(
                Path.GetDirectoryName(typeof(SplitWallHandler).Assembly.Location)!,
                "split_results.json");

        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Get selected walls
                var walls = WallSelectionHelper.GetSelectedWalls(uidoc);
                if (walls.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "No walls selected. Please select one or more walls.");
                    return;
                }

                // Filter to compound walls only
                var compoundWalls = walls
                    .Where(w => CompoundLayerAnalyzer.IsCompoundWall(w))
                    .ToList();

                if (compoundWalls.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "None of the selected walls are compound walls with multiple layers.");
                    return;
                }

                // Use the first compound wall for dialog setup
                var firstWall = compoundWalls[0];
                var layers = CompoundLayerAnalyzer.AnalyzeLayers(firstWall);

                // Build layer rows for the dialog
                var layerRows = new List<LayerRowItem>();
                foreach (var layer in layers)
                {
                    double thicknessMm = Math.Round(
                        UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters), 1);

                    // Find existing types matching this layer's thickness (read-only, no transaction)
                    var matchingTypes = WallTypeResolver.FindMatchingTypes(doc, layer.Thickness);

                    var availableTypes = new List<WallTypeOption>();

                    // Add auto-create option first
                    availableTypes.Add(new WallTypeOption
                    {
                        DisplayName = $"* Auto-create ({layer.MaterialName}, {thicknessMm:F1}mm)",
                        WallType = null,
                        IsAutoCreate = true
                    });

                    // Add existing matching types
                    foreach (var wt in matchingTypes)
                    {
                        availableTypes.Add(new WallTypeOption
                        {
                            DisplayName = wt.Name,
                            WallType = wt,
                            IsAutoCreate = false
                        });
                    }

                    var row = new LayerRowItem
                    {
                        Index = layer.Index,
                        FunctionName = layer.Function.ToString(),
                        MaterialName = layer.MaterialName,
                        ThicknessMm = $"{thicknessMm:F1}",
                        AvailableTypes = availableTypes,
                        SelectedType = availableTypes[0] // Default to auto-create
                    };

                    layerRows.Add(row);
                }

                // Count hosted elements on the first wall
                int hostedCount = HostedElementTransfer.GetHostedElements(doc, firstWall).Count;

                // Compute total width in mm
                double widthMm = Math.Round(
                    UnitUtils.ConvertFromInternalUnits(firstWall.WallType.Width, UnitTypeId.Millimeters), 1);

                // Show dialog
                var dialog = new WallSplitterDialog(
                    firstWall.WallType.Name, widthMm, layerRows, hostedCount);
                dialog.ShowDialog();

                if (!dialog.Accepted)
                    return;

                // Map dialog selections back to LayerInfo.ResolvedType
                foreach (var layer in layers)
                {
                    var row = layerRows.FirstOrDefault(r => r.Index == layer.Index);
                    if (row?.SelectedType != null && !row.SelectedType.IsAutoCreate)
                    {
                        layer.ResolvedType = row.SelectedType.WallType;
                    }
                }

                // Determine scope
                List<Wall> wallsToSplit;
                if (dialog.SplitAllOfType)
                {
                    var wallTypeId = firstWall.WallType.Id;
                    wallsToSplit = new FilteredElementCollector(doc)
                        .OfClass(typeof(Wall))
                        .Cast<Wall>()
                        .Where(w => w.WallType.Id == wallTypeId)
                        .Where(w => WallSelectionHelper.IsValidWall(w))
                        .Where(w => CompoundLayerAnalyzer.IsCompoundWall(w))
                        .ToList();
                }
                else
                {
                    wallsToSplit = compoundWalls;
                }

                // Execute splits
                int succeeded = 0;
                int failed = 0;
                var messages = new List<string>();
                var allResults = new List<SplitResult>();
                var wallIds = wallsToSplit.Select(w => w.Id).ToList();

                foreach (var wallId in wallIds)
                {
                    var wall = doc.GetElement(wallId) as Wall;
                    if (wall == null || !wall.IsValidObject)
                    {
                        failed++;
                        messages.Add($"Wall {wallId}: Element no longer valid.");
                        continue;
                    }

                    var wallLayers = CompoundLayerAnalyzer.AnalyzeLayers(wall);
                    if (wallLayers.Count == 0)
                    {
                        failed++;
                        messages.Add($"Wall {wall.Id}: No splittable layers.");
                        continue;
                    }

                    foreach (var wl in wallLayers)
                    {
                        var row = layerRows.FirstOrDefault(r => r.Index == wl.Index);
                        if (row?.SelectedType != null && !row.SelectedType.IsAutoCreate)
                            wl.ResolvedType = row.SelectedType.WallType;
                    }

                    var result = WallSplitterEngine.SplitWall(
                        doc, wall, wallLayers, dialog.SelectedHostLayerIndex);

                    allResults.Add(result);

                    if (result.Success)
                        succeeded++;
                    else
                    {
                        failed++;
                        messages.Add($"Wall {wall.Id}: {result.Message}");
                    }
                }

                // Save split results to JSON for the Trim Corners command
                var successfulResults = allResults.Where(r => r.Success).ToList();
                SaveSplitResults(successfulResults);

                // Show summary
                string summary = $"Split complete: {succeeded} succeeded, {failed} failed out of {wallsToSplit.Count} wall(s).";
                if (messages.Count > 0)
                    summary += "\n\nDetails:\n" + string.Join("\n", messages);

                summary += "\n\nOK any Revit warnings, then click 'Trim Corners' in the panel.";
                TaskDialog.Show("BimTasksV2 - Wall Splitter", summary);
                Log.Information("[SplitWallHandler] {Summary}", summary.Replace("\n", " "));

                // Show the dockable panel with TrimCorners view
                try
                {
                    var pane = uiApp.GetDockablePane(
                        BimTasksV2.Infrastructure.BimTasksBootstrapper.DockablePaneId);
                    if (pane != null && !pane.IsShown())
                        pane.Show();

                    var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                    eventAgg.GetEvent<Events.BimTasksEvents.SwitchDockablePanelEvent>()
                        .Publish("TrimCorners");
                }
                catch (Exception paneEx)
                {
                    Log.Warning(paneEx, "[SplitWallHandler] Could not show dockable panel");
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[SplitWallHandler] Failed. StackTrace: {Stack}", ex.StackTrace);
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}\n\nSource: {ex.TargetSite?.DeclaringType?.Name}.{ex.TargetSite?.Name}");
            }
        }

        /// <summary>
        /// Serializes split result IDs and layer info to a JSON file
        /// so the Trim Corners command can read them as a separate invocation.
        /// </summary>
        private static void SaveSplitResults(List<SplitResult> results)
        {
            try
            {
                var data = results.Select(r => new SplitResultData
                {
                    Replacements = r.ReplacementIds.Select(rid => new ReplacementData
                    {
                        WallId = rid.WallId.Value,
                        LayerIndex = rid.Layer.Index
                    }).ToList()
                }).ToList();

                var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SplitResultsPath, json);
                Log.Information("[SplitWallHandler] Saved {Count} split results to {Path}", results.Count, SplitResultsPath);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[SplitWallHandler] Failed to save split results");
            }
        }
    }

    /// <summary>JSON-serializable split result data.</summary>
    public class SplitResultData
    {
        public List<ReplacementData> Replacements { get; set; } = new();
    }

    /// <summary>JSON-serializable replacement wall reference.</summary>
    public class ReplacementData
    {
        public long WallId { get; set; }
        public int LayerIndex { get; set; }
    }
}
