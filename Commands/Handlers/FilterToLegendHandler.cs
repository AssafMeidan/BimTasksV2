using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Helpers;
using BimTasksV2.Views;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    public class FilterToLegendHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                var view = doc.ActiveView;

                // Collect filters applied to the active view with their overrides
                var filterItems = CollectViewFilters(doc, view);

                if (filterItems.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "No filters with color overrides are applied to the active view.");
                    return;
                }

                // Collect existing legend views for the update option
                // List both Legend and Drafting views (our legends are created as drafting views)
                // Filter to those starting with "Legend -" for drafting views
                var existingLegends = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => !v.IsTemplate &&
                        (v.ViewType == ViewType.Legend ||
                         (v.ViewType == ViewType.DraftingView && v.Name.StartsWith("Legend -"))))
                    .OrderBy(v => v.Name)
                    .Select(v => new Views.LegendViewInfo
                    {
                        ViewId = v.Id.Value,
                        Name = v.Name
                    })
                    .ToList();

                // Show checklist dialog
                var dialog = new FilterToLegendDialog(filterItems, existingLegends);
                dialog.ShowDialog();

                if (!dialog.Accepted || dialog.SelectedFilters.Count == 0)
                    return;

                var legendBuilder = new LegendBuilder();

                if (dialog.IsCreateNew)
                {
                    // Create new legend
                    using (var tg = new TransactionGroup(doc, "Create Filter Legend"))
                    {
                        tg.Start();
                        var legendView = legendBuilder.CreateLegend(doc, view.Name, dialog.SelectedFilters);
                        tg.Assimilate();

                        if (legendView != null)
                        {
                            uidoc.ActiveView = legendView;
                            Log.Information("[FilterToLegend] Created legend '{Name}' with {Count} items",
                                legendView.Name, dialog.SelectedFilters.Count);
                        }
                    }
                }
                else
                {
                    // Update existing legend
                    var targetViewId = new ElementId(dialog.SelectedExistingLegend!.ViewId);
                    var targetView = doc.GetElement(targetViewId) as View;

                    if (targetView == null)
                    {
                        TaskDialog.Show("BimTasksV2", "Selected legend view no longer exists.");
                        return;
                    }

                    using (var tg = new TransactionGroup(doc, "Update Filter Legend"))
                    {
                        tg.Start();
                        legendBuilder.UpdateLegend(doc, targetView, dialog.SelectedFilters);
                        tg.Assimilate();

                        uidoc.ActiveView = targetView;
                        Log.Information("[FilterToLegend] Updated legend '{Name}' with {Count} items",
                            targetView.Name, dialog.SelectedFilters.Count);
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // User cancelled
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[FilterToLegend] Failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }

        private static List<FilterLegendItem> CollectViewFilters(Document doc, View view)
        {
            var items = new List<FilterLegendItem>();
            var filterIds = view.GetFilters();

            foreach (var filterId in filterIds)
            {
                var filter = doc.GetElement(filterId) as ParameterFilterElement;
                if (filter == null) continue;

                var overrides = view.GetFilterOverrides(filterId);

                // Try surface foreground color first, then cut foreground as fallback
                var color = overrides.SurfaceForegroundPatternColor;
                var patternId = overrides.SurfaceForegroundPatternId;

                if (!color.IsValid || patternId == ElementId.InvalidElementId)
                {
                    var cutColor = overrides.CutForegroundPatternColor;
                    var cutPatternId = overrides.CutForegroundPatternId;
                    if (cutColor.IsValid && cutPatternId != ElementId.InvalidElementId)
                    {
                        color = cutColor;
                        patternId = cutPatternId;
                    }
                }

                // Only include filters that have a valid color override
                if (!color.IsValid)
                    continue;

                items.Add(new FilterLegendItem
                {
                    FilterName = filter.Name,
                    FilterId = filterId,
                    OverrideColor = color,
                    FillPatternId = patternId,
                    HasColorOverride = true
                });
            }

            return items;
        }

    }

    public class FilterLegendItem
    {
        public string FilterName { get; set; } = "";
        public ElementId FilterId { get; set; } = ElementId.InvalidElementId;
        public Color OverrideColor { get; set; } = new Color(200, 200, 200);
        public ElementId FillPatternId { get; set; } = ElementId.InvalidElementId;
        public bool HasColorOverride { get; set; }
        public bool IsSelected { get; set; } = true;
    }
}
