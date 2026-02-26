using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Events;
using BimTasksV2.Helpers;
using BimTasksV2.Helpers.WallSplitter;
using BimTasksV2.Views;
using Prism.Events;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Splits compound walls into individual single-layer walls.
    /// Shows a dialog for layer type assignment, hosted element handling, and scope selection.
    /// </summary>
    public class SplitWallHandler : ICommandHandler
    {
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
                    // If auto-create or no selection, leave ResolvedType null — engine will auto-create
                }

                // Determine scope
                List<Wall> wallsToSplit;
                if (dialog.SplitAllOfType)
                {
                    // Collect all walls of the same WallType in the project
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

                foreach (var wall in wallsToSplit)
                {
                    // Re-analyze layers per wall (offsets are per-instance)
                    var wallLayers = CompoundLayerAnalyzer.AnalyzeLayers(wall);
                    if (wallLayers.Count == 0)
                    {
                        failed++;
                        messages.Add($"Wall {wall.Id}: No splittable layers.");
                        continue;
                    }

                    // Apply type selections from dialog to re-analyzed layers
                    foreach (var wl in wallLayers)
                    {
                        var row = layerRows.FirstOrDefault(r => r.Index == wl.Index);
                        if (row?.SelectedType != null && !row.SelectedType.IsAutoCreate)
                        {
                            wl.ResolvedType = row.SelectedType.WallType;
                        }
                    }

                    var result = WallSplitterEngine.SplitWall(
                        doc, wall, wallLayers, dialog.SelectedHostLayerIndex);

                    allResults.Add(result);

                    if (result.Success)
                    {
                        succeeded++;
                    }
                    else
                    {
                        failed++;
                        messages.Add($"Wall {wall.Id}: {result.Message}");
                    }
                }

                // Save corner fix data and show Fix Corners panel
                var successfulResults = allResults.Where(r => r.Success).ToList();
                if (successfulResults.Count > 0)
                {
                    try
                    {
                        string filePath = WallSplitterEngine.GetCornerFixDataPath();
                        WallSplitterEngine.SaveCornerFixData(filePath, successfulResults);
                        Log.Information("[SplitWallHandler] Saved corner fix data to {FilePath}", filePath);
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "[SplitWallHandler] Failed to save corner fix data");
                    }

                    // Switch dockable panel to FixSplitCornersView
                    try
                    {
                        var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                        var contextService = container.Resolve<Services.IRevitContextService>();
                        contextService.UIApplication = uiApp;
                        contextService.UIDocument = uidoc;

                        var eventAggregator = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                        eventAggregator.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                            .Publish("FixSplitCorners");

                        int totalReplacements = successfulResults.Sum(r => r.ReplacementWalls.Count);
                        eventAggregator.GetEvent<BimTasksEvents.FixSplitCornersReadyEvent>()
                            .Publish(new FixSplitCornersPayload
                            {
                                WallsSplit = succeeded,
                                TotalReplacements = totalReplacements
                            });

                        var pane = uiApp.GetDockablePane(
                            BimTasksV2.Infrastructure.BimTasksBootstrapper.DockablePaneId);
                        if (pane != null && !pane.IsShown())
                            pane.Show();
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "[SplitWallHandler] Failed to show Fix Corners panel");
                    }
                }

                // Show summary
                string summary = $"Split complete: {succeeded} succeeded, {failed} failed out of {wallsToSplit.Count} wall(s).";
                if (successfulResults.Count > 0)
                {
                    summary += "\n\nDismiss any join errors, then click 'Fix Corners' in the BimTasks panel.";
                }
                if (messages.Count > 0)
                {
                    summary += "\n\nDetails:\n" + string.Join("\n", messages);
                }
                TaskDialog.Show("BimTasksV2 - Wall Splitter", summary);

                Log.Information("[SplitWallHandler] {Summary}", summary.Replace("\n", " "));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled — silently return
            }
            catch (Exception ex)
            {
                Log.Error(ex, "SplitWall failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
