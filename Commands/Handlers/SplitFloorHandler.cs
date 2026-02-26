using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers.FloorSplitter;
using BimTasksV2.Views;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Splits compound floors into individual single-layer floors.
    /// Shows a dialog for layer type assignment and scope selection.
    /// </summary>
    public class SplitFloorHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Get selected floors
                var selectedIds = uidoc.Selection.GetElementIds();
                var floors = selectedIds
                    .Select(id => doc.GetElement(id))
                    .OfType<Floor>()
                    .ToList();

                if (floors.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "No floors selected. Please select one or more floors.");
                    return;
                }

                // Filter to compound floors only
                var compoundFloors = floors
                    .Where(f => FloorLayerAnalyzer.IsCompoundFloor(f))
                    .ToList();

                if (compoundFloors.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "None of the selected floors are compound floors with multiple layers.");
                    return;
                }

                // Use the first compound floor for dialog setup
                var firstFloor = compoundFloors[0];
                var firstFloorType = doc.GetElement(firstFloor.GetTypeId()) as FloorType;
                var layers = FloorLayerAnalyzer.AnalyzeLayers(firstFloor);

                // Compute total thickness
                var cs = firstFloorType?.GetCompoundStructure();
                double totalThickness = cs?.GetLayers().Sum(l => l.Width) ?? 0;
                double thicknessMm = Math.Round(
                    UnitUtils.ConvertFromInternalUnits(totalThickness, UnitTypeId.Millimeters), 1);

                // Build layer rows for the dialog
                var layerRows = new List<FloorLayerRowItem>();
                foreach (var layer in layers)
                {
                    double layerMm = Math.Round(
                        UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters), 1);

                    var matchingTypes = FloorTypeResolver.FindMatchingTypes(doc, layer.Thickness);
                    var availableTypes = new List<FloorTypeOption>();

                    availableTypes.Add(new FloorTypeOption
                    {
                        DisplayName = $"* Auto-create ({layer.MaterialName}, {layerMm:F1}mm)",
                        FloorType = null,
                        IsAutoCreate = true
                    });

                    foreach (var ft in matchingTypes)
                    {
                        availableTypes.Add(new FloorTypeOption
                        {
                            DisplayName = ft.Name,
                            FloorType = ft,
                            IsAutoCreate = false
                        });
                    }

                    layerRows.Add(new FloorLayerRowItem
                    {
                        Index = layer.Index,
                        FunctionName = layer.Function.ToString(),
                        MaterialName = layer.MaterialName,
                        ThicknessMm = $"{layerMm:F1}",
                        AvailableTypes = availableTypes,
                        SelectedType = availableTypes[0]
                    });
                }

                // Show dialog
                var dialog = new FloorSplitterDialog(
                    firstFloorType?.Name ?? "Unknown", thicknessMm, layerRows);
                dialog.ShowDialog();

                if (!dialog.Accepted)
                    return;

                // Map dialog selections back to layer info
                foreach (var layer in layers)
                {
                    var row = layerRows.FirstOrDefault(r => r.Index == layer.Index);
                    if (row?.SelectedType != null && !row.SelectedType.IsAutoCreate)
                        layer.ResolvedType = row.SelectedType.FloorType;
                }

                // Determine scope
                List<Floor> floorsToSplit;
                if (dialog.SplitAllOfType)
                {
                    var floorTypeId = firstFloor.GetTypeId();
                    floorsToSplit = new FilteredElementCollector(doc)
                        .OfClass(typeof(Floor))
                        .Cast<Floor>()
                        .Where(f => f.GetTypeId() == floorTypeId)
                        .Where(f => FloorLayerAnalyzer.IsCompoundFloor(f))
                        .ToList();
                }
                else
                {
                    floorsToSplit = compoundFloors;
                }

                // Execute splits
                int succeeded = 0;
                int failed = 0;
                var messages = new List<string>();

                foreach (var floor in floorsToSplit)
                {
                    // Re-analyze layers per floor
                    var floorLayers = FloorLayerAnalyzer.AnalyzeLayers(floor);
                    if (floorLayers.Count == 0)
                    {
                        failed++;
                        messages.Add($"Floor {floor.Id}: No splittable layers.");
                        continue;
                    }

                    // Apply type selections from dialog
                    foreach (var fl in floorLayers)
                    {
                        var row = layerRows.FirstOrDefault(r => r.Index == fl.Index);
                        if (row?.SelectedType != null && !row.SelectedType.IsAutoCreate)
                            fl.ResolvedType = row.SelectedType.FloorType;
                    }

                    var result = FloorSplitterEngine.SplitFloor(doc, floor, floorLayers);

                    if (result.Success)
                        succeeded++;
                    else
                    {
                        failed++;
                        messages.Add($"Floor {floor.Id}: {result.Message}");
                    }
                }

                // Show summary
                string summary = $"Split complete: {succeeded} succeeded, {failed} failed out of {floorsToSplit.Count} floor(s).";
                if (messages.Count > 0)
                    summary += "\n\nDetails:\n" + string.Join("\n", messages);

                TaskDialog.Show("BimTasksV2 - Floor Splitter", summary);
                Log.Information("[SplitFloorHandler] {Summary}", summary.Replace("\n", " "));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "SplitFloor failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
