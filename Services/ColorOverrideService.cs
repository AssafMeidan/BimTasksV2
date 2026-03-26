using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Serilog;

namespace BimTasksV2.Services
{
    public class ColorSwatchGroup
    {
        public string Value { get; set; } = "";
        public string Description { get; set; } = "";
        public System.Windows.Media.Color Color { get; set; }
        public List<ElementId> ElementIds { get; set; } = new();
    }

    public class ColorOverrideService
    {
        private readonly Dictionary<ElementId, OverrideGraphicSettings> _originalOverrides = new();
        private ElementId _solidFillPatternId = ElementId.InvalidElementId;

        public bool HasActiveOverrides => _originalOverrides.Count > 0;

        /// <summary>
        /// Scans visible elements in the view for unique values of the given shared parameter.
        /// Groups elements by value and optionally reads a secondary parameter for description.
        /// </summary>
        public List<ColorSwatchGroup> ScanElements(
            Document doc,
            View view,
            string parameterName,
            string? descriptionParameterName,
            IReadOnlyCollection<BuiltInCategory> categories)
        {
            var groups = new Dictionary<string, ColorSwatchGroup>();

            foreach (var category in categories)
            {
                var collector = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                foreach (var element in collector)
                {
                    var param = element.LookupParameter(parameterName);
                    if (param == null) continue;

                    string value = GetParameterValueAsString(param);
                    string key = value ?? "";

                    if (!groups.TryGetValue(key, out var group))
                    {
                        group = new ColorSwatchGroup
                        {
                            Value = string.IsNullOrEmpty(key) ? "(empty)" : key,
                        };

                        // Read description from the first element found
                        if (!string.IsNullOrEmpty(descriptionParameterName))
                        {
                            var descParam = element.LookupParameter(descriptionParameterName);
                            if (descParam != null)
                                group.Description = GetParameterValueAsString(descParam) ?? "";
                        }

                        groups[key] = group;
                    }

                    group.ElementIds.Add(element.Id);
                }
            }

            return groups.Values
                .OrderByDescending(g => g.Value == "(empty)" ? 0 : 1)
                .ThenBy(g => g.Value)
                .ToList();
        }

        /// <summary>
        /// Collects all shared parameter names found on visible model elements.
        /// </summary>
        public List<string> GetSharedParameterNames(
            Document doc,
            View view,
            IReadOnlyCollection<BuiltInCategory> categories)
        {
            var names = new HashSet<string>();

            foreach (var category in categories)
            {
                var collector = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                // Sample up to 50 elements per category for performance
                foreach (var element in collector.Take(50))
                {
                    foreach (Parameter param in element.Parameters)
                    {
                        if (param.IsShared && !string.IsNullOrEmpty(param.Definition?.Name))
                            names.Add(param.Definition.Name);
                    }
                }
            }

            return names.OrderBy(n => n).ToList();
        }

        /// <summary>
        /// Finds the first shared parameter (from the given list) that has a non-empty value
        /// on at least one visible element. Used to auto-select a default description parameter.
        /// </summary>
        public string? FindFirstParameterWithValue(
            Document doc,
            View view,
            IReadOnlyCollection<BuiltInCategory> categories,
            IEnumerable<string> parameterNames,
            string? excludeName)
        {
            foreach (var name in parameterNames)
            {
                if (name == excludeName) continue;

                foreach (var category in categories)
                {
                    var collector = new FilteredElementCollector(doc, view.Id)
                        .OfCategory(category)
                        .WhereElementIsNotElementType();

                    foreach (var element in collector.Take(20))
                    {
                        var param = element.LookupParameter(name);
                        if (param == null) continue;

                        var val = GetParameterValueAsString(param);
                        if (!string.IsNullOrEmpty(val) && val.Any(c => char.IsLetterOrDigit(c)))
                            return name;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Gets model categories with element counts in the view, excluding annotations.
        /// Returns sorted by count descending.
        /// </summary>
        public List<(BuiltInCategory Category, string Name, int Count)> GetModelCategories(
            Document doc, View view)
        {
            var result = new List<(BuiltInCategory, string, int)>();

            var allElements = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .ToElements();

            var grouped = allElements
                .Where(e => e.Category != null
                    && e.Category.CategoryType == CategoryType.Model
                    && (BuiltInCategory)e.Category.Id.Value != BuiltInCategory.INVALID)
                .GroupBy(e => (BuiltInCategory)e.Category.Id.Value)
                .Select(g => (
                    Category: g.Key,
                    Name: g.First().Category.Name,
                    Count: g.Count()))
                .OrderByDescending(x => x.Count)
                .ToList();

            return grouped;
        }

        /// <summary>
        /// Applies color overrides to elements in the view. Stores originals for restoration.
        /// </summary>
        public void ApplyOverrides(
            Document doc,
            View view,
            List<ColorSwatchGroup> groups)
        {
            EnsureSolidFillPattern(doc);

            using var tx = new Transaction(doc, "Color Swatch — Apply");
            tx.Start();

            try
            {
                foreach (var group in groups)
                {
                    var revitColor = new Color(group.Color.R, group.Color.G, group.Color.B);

                    foreach (var elementId in group.ElementIds)
                    {
                        // Store original override (only if not already stored)
                        if (!_originalOverrides.ContainsKey(elementId))
                        {
                            var original = view.GetElementOverrides(elementId);
                            _originalOverrides[elementId] = original;
                        }

                        var ogs = new OverrideGraphicSettings();
                        ogs.SetSurfaceForegroundPatternId(_solidFillPatternId);
                        ogs.SetSurfaceForegroundPatternColor(revitColor);
                        view.SetElementOverrides(elementId, ogs);
                    }
                }

                tx.Commit();
                Log.Information("[ColorOverride] Applied overrides to {Count} elements",
                    _originalOverrides.Count);
            }
            catch (Exception ex)
            {
                if (tx.HasStarted()) tx.RollBack();
                Log.Error(ex, "[ColorOverride] Failed to apply overrides");
                throw;
            }
        }

        /// <summary>
        /// Scans all visible elements in the view and resets any that have
        /// solid-fill surface foreground pattern overrides (i.e., overrides
        /// applied by the Color Swatch tool). Works even for overrides from
        /// previous sessions that are no longer tracked in memory.
        /// </summary>
        public int ClearAllSwatchOverrides(Document doc, View view)
        {
            EnsureSolidFillPattern(doc);
            if (_solidFillPatternId == ElementId.InvalidElementId) return 0;

            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();

            using var tx = new Transaction(doc, "Color Swatch — Reset View");
            tx.Start();

            int count = 0;
            try
            {
                foreach (var element in collector)
                {
                    var ogs = view.GetElementOverrides(element.Id);
                    if (ogs.SurfaceForegroundPatternId == _solidFillPatternId)
                    {
                        view.SetElementOverrides(element.Id, new OverrideGraphicSettings());
                        count++;
                    }
                }

                tx.Commit();
                Log.Information("[ColorOverride] Reset {Count} element overrides in view '{View}'",
                    count, view.Name);
            }
            catch (Exception ex)
            {
                if (tx.HasStarted()) tx.RollBack();
                Log.Error(ex, "[ColorOverride] Failed to reset view overrides");
                throw;
            }

            // Also remove any tracked originals for this view's elements
            // so Clear Colors doesn't try to restore stale data
            _originalOverrides.Clear();

            return count;
        }

        /// <summary>
        /// Restores all original overrides and clears state.
        /// </summary>
        public void ClearOverrides(Document doc, View view)
        {
            if (_originalOverrides.Count == 0) return;

            using var tx = new Transaction(doc, "Color Swatch — Clear");
            tx.Start();

            try
            {
                foreach (var kvp in _originalOverrides)
                {
                    if (doc.GetElement(kvp.Key) != null)
                        view.SetElementOverrides(kvp.Key, kvp.Value);
                }

                tx.Commit();
                Log.Information("[ColorOverride] Cleared overrides for {Count} elements",
                    _originalOverrides.Count);
            }
            catch (Exception ex)
            {
                if (tx.HasStarted()) tx.RollBack();
                Log.Error(ex, "[ColorOverride] Failed to clear overrides");
                throw;
            }

            _originalOverrides.Clear();
        }

        private void EnsureSolidFillPattern(Document doc)
        {
            if (_solidFillPatternId != ElementId.InvalidElementId) return;

            var solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            _solidFillPatternId = solidFill?.Id ?? ElementId.InvalidElementId;

            if (_solidFillPatternId == ElementId.InvalidElementId)
                Log.Warning("[ColorOverride] No solid fill pattern found in document");
        }

        private static string? GetParameterValueAsString(Parameter param)
        {
            return param.StorageType switch
            {
                StorageType.String => param.AsString(),
                StorageType.Integer => param.AsInteger().ToString(),
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString("F2"),
                StorageType.ElementId => param.AsValueString() ?? param.AsElementId()?.ToString(),
                _ => null
            };
        }
    }
}
