using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Events;
using BimTasksV2.Helpers;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class FilterTreeViewModel : BindableBase
    {
        #region Fields

        private readonly IRevitContextService _revitContext;

        private static readonly HashSet<BuiltInParameter> IncludedBuiltInParameters =
            CategoryFilterHelper.GetBuiltParamsToInclude();

        private static readonly HashSet<BuiltInCategory> CategoriesToInclude =
            new(CategoryFilterHelper.GetCategoriesToInclude());

        /// <summary>Store ElementIds (not Element refs) to avoid stale references.</summary>
        private List<ElementId> _elementIdsToProcess = new();

        #endregion

        #region Properties

        public ObservableCollection<TreeNodeViewModel> ElementTypes { get; } = new();

        private string _topHeader = "Filter Tree";
        public string TopHeader
        {
            get => _topHeader;
            set => SetProperty(ref _topHeader, value);
        }

        private bool _isAndMode = true;
        public bool IsAndMode
        {
            get => _isAndMode;
            set => SetProperty(ref _isAndMode, value);
        }

        #endregion

        #region Commands

        public DelegateCommand OKCommand { get; }
        public DelegateCommand ResetCommand { get; }

        #endregion

        #region Constructor

        public FilterTreeViewModel()
        {
            _revitContext = BimTasksV2.Infrastructure.ContainerLocator.Container
                .Resolve<IRevitContextService>();

            OKCommand = new DelegateCommand(OnOK);
            ResetCommand = new DelegateCommand(OnReset);

            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.ResetFilterTreeEvent>()
                .Subscribe(_ => SafeLoadTree(), ThreadOption.PublisherThread, keepSubscriberReferenceAlive: false);
        }

        #endregion

        #region Tree Loading

        private void SafeLoadTree()
        {
            try
            {
                var uiDoc = _revitContext.UIDocument;
                if (uiDoc == null)
                {
                    Log.Warning("[FilterTree] No active document");
                    return;
                }

                var currentSelection = uiDoc.Selection.GetElementIds();
                LoadElementTypes(uiDoc, currentSelection.Count > 0 ? currentSelection : null);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[FilterTree] Error loading tree");
                TaskDialog.Show("Filter Tree Error", $"Failed to load elements:\n{ex.Message}");
            }
        }

        private void LoadElementTypes(UIDocument uiDoc, ICollection<ElementId>? selectedIds)
        {
            var doc = uiDoc.Document;
            Units projectUnits;
            try { projectUnits = doc.GetUnits(); }
            catch
            {
                Log.Warning("[FilterTree] Could not get document units");
                return;
            }

            ElementTypes.Clear();
            _elementIdsToProcess.Clear();

            IList<Element> elements;

            if (selectedIds != null && selectedIds.Count > 0)
            {
                var validIds = selectedIds
                    .Where(id => id != null && id != ElementId.InvalidElementId)
                    .Where(id => doc.GetElement(id) != null)
                    .ToList();

                if (validIds.Count == 0) return;

                elements = new FilteredElementCollector(doc, validIds)
                    .WhereElementIsNotElementType()
                    .Where(e => IsValidElement(e))
                    .ToList();
            }
            else
            {
                elements = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => IsValidElement(e))
                    .ToList();
            }

            _elementIdsToProcess = elements.Select(e => e.Id).ToList();
            TopHeader = selectedIds != null
                ? $"Filter ({elements.Count} selected)"
                : $"Filter ({elements.Count} elements)";

            if (elements.Count == 0) return;

            // Group by category name
            var byCategory = elements
                .Where(e => e?.Category != null)
                .GroupBy(e => e.Category.Name)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase);

            foreach (var catGroup in byCategory)
            {
                try
                {
                    var catNode = BuildCategoryNode(catGroup, doc, projectUnits);
                    if (catNode != null && catNode.Children.Count > 0)
                        ElementTypes.Add(catNode);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[FilterTree] Error processing category {Cat}", catGroup.Key);
                }
            }
        }

        private static bool IsValidElement(Element e)
        {
            if (e?.Category == null) return false;
            try
            {
                return CategoriesToInclude.Contains((BuiltInCategory)e.Category.Id.Value);
            }
            catch { return false; }
        }

        private TreeNodeViewModel? BuildCategoryNode(
            IGrouping<string, Element> catGroup, Document doc, Units projectUnits)
        {
            var catName = catGroup.Key;
            var first = catGroup.FirstOrDefault();
            if (first?.Category == null) return null;

            var catNode = new TreeNodeViewModel(catName)
            {
                ElementId = first.Category.Id,
                DisplayName = $"{catName} ({catGroup.Count()})"
            };

            var elementsInCat = catGroup.ToList();

            // Collect all valid parameter names across elements in this category
            var paramNames = elementsInCat
                .SelectMany(e => GetValidParameterNames(e))
                .Distinct()
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase);

            foreach (var paramName in paramNames)
            {
                try
                {
                    var propNode = BuildPropertyNode(paramName, elementsInCat, catNode, doc, projectUnits);
                    if (propNode != null && propNode.Children.Count > 0)
                        catNode.Children.Add(propNode);
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "[FilterTree] Error building param node {Param}", paramName);
                }
            }

            return catNode;
        }

        private IEnumerable<string> GetValidParameterNames(Element e)
        {
            if (e?.Parameters == null)
                return Enumerable.Empty<string>();

            try
            {
                return e.Parameters
                    .OfType<Parameter>()
                    .Where(p => p?.Definition?.Name != null && !IsExcludedParameter(p))
                    .Select(p => p.Definition.Name);
            }
            catch { return Enumerable.Empty<string>(); }
        }

        private TreeNodeViewModel BuildPropertyNode(
            string paramName, List<Element> elements,
            TreeNodeViewModel catNode, Document doc, Units projectUnits)
        {
            var propNode = new TreeNodeViewModel(paramName, catNode);

            var valueGroups = elements
                .Select(e =>
                {
                    try
                    {
                        var p = e.LookupParameter(paramName);
                        return p != null
                            ? new { Value = GetParameterValueSafe(p, doc, projectUnits), Element = e }
                            : null;
                    }
                    catch { return null; }
                })
                .Where(x => x != null && !string.IsNullOrEmpty(x!.Value))
                .GroupBy(x => x!.Value)
                .OrderBy(g => g.Key ?? "", StringComparer.OrdinalIgnoreCase);

            foreach (var vg in valueGroups)
            {
                string val = vg.Key ?? "(empty)";
                var valueNode = new TreeNodeViewModel(val, propNode)
                {
                    PropertyName = paramName,
                    PropertyValue = vg.Key,
                    DisplayName = $"{val} ({vg.Count()})"
                };
                propNode.Children.Add(valueNode);
            }

            return propNode;
        }

        #endregion

        #region Parameter Helpers

        private bool IsExcludedParameter(Parameter param)
        {
            if (param?.Definition == null) return true;

            try
            {
                // Allow known BuiltInParameters (e.g. WALL_USER_HEIGHT_PARAM)
                if (!param.IsShared && Enum.IsDefined(typeof(BuiltInParameter), param.Id.Value))
                {
                    var bip = (BuiltInParameter)param.Id.Value;
                    if (IncludedBuiltInParameters.Contains(bip))
                        return false;
                }

                string name = param.Definition.Name ?? "";

                // Exclude GUID and ID parameters
                if (name.Contains("GUID", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    return true;

                // Exclude Double (dimension) parameters unless they're in the BuiltIn list
                if (param.StorageType == StorageType.Double)
                    return true;

                return false;
            }
            catch { return true; }
        }

        private static string GetParameterValueSafe(Parameter param, Document doc, Units projectUnits)
        {
            if (param == null) return "";

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? "";

                    case StorageType.Double:
                        try
                        {
                            var opts = projectUnits.GetFormatOptions(SpecTypeId.Length);
                            var unitId = opts.GetUnitTypeId();
                            double val = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), unitId);
                            return val.ToString("F2");
                        }
                        catch
                        {
                            return param.AsDouble().ToString("F2");
                        }

                    case StorageType.Integer:
                        return param.AsInteger().ToString();

                    case StorageType.ElementId:
                        var id = param.AsElementId();
                        if (id == null || id == ElementId.InvalidElementId) return "";
                        try
                        {
                            var elem = doc.GetElement(id);
                            return elem?.Name ?? id.Value.ToString();
                        }
                        catch { return id.Value.ToString(); }

                    default:
                        return "";
                }
            }
            catch { return ""; }
        }

        #endregion

        #region Commands Implementation

        private void OnOK()
        {
            try
            {
                var selected = GetSelectedItems();
                if (selected.Count == 0)
                {
                    TaskDialog.Show("Filter", "No items selected. Please check items in the tree.");
                    return;
                }

                SelectElementsInRevit(selected);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[FilterTree] OK failed");
                TaskDialog.Show("Filter Error", $"Failed to select elements:\n{ex.Message}");
            }
        }

        private void OnReset()
        {
            try
            {
                // Apply current selection first
                var selected = GetSelectedItems();
                if (selected.Count > 0)
                    SelectElementsInRevit(selected);

                // Then reload tree with current Revit selection
                SafeLoadTree();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[FilterTree] Reset failed");
                TaskDialog.Show("Filter Error", $"Failed to reset filter:\n{ex.Message}");
            }
        }

        #endregion

        #region Selection Logic

        private List<SelectedItem> GetSelectedItems()
        {
            var items = new List<SelectedItem>();
            foreach (var node in ElementTypes)
                CollectSelectedItems(node, items);
            return items;
        }

        private void CollectSelectedItems(TreeNodeViewModel node, List<SelectedItem> items)
        {
            if (node.IsChecked && node.ElementId != null)
            {
                // Category-level node is fully checked — select by category, not by params
                items.Add(new SelectedItem
                {
                    CategoryName = node.Name,
                    CategoryId = node.ElementId,
                    IsCategoryMatch = true
                });
                return;
            }

            if (node.IsChecked)
            {
                if (node.Children.Count > 0)
                {
                    // Has children — recurse into them
                    foreach (var child in node.Children)
                        CollectSelectedItems(child, items);
                }
                else
                {
                    // Leaf node (value node) — collect it
                    items.Add(new SelectedItem
                    {
                        CategoryName = GetParentCategoryName(node),
                        CategoryId = GetParentCategoryId(node),
                        PropertyName = node.PropertyName ?? "",
                        PropertyValue = node.PropertyValue ?? node.Name
                    });
                }
            }
            else if (node.Children.Count > 0)
            {
                // Not checked, but children might be individually checked
                foreach (var child in node.Children)
                    CollectSelectedItems(child, items);
            }
        }

        private void SelectElementsInRevit(List<SelectedItem> selectedItems)
        {
            var uiDoc = _revitContext.UIDocument;
            if (uiDoc == null) return;
            var doc = uiDoc.Document;

            try
            {
                Units projectUnits = doc.GetUnits();

                // Get fresh elements from stored IDs
                var validElements = _elementIdsToProcess
                    .Select(id => { try { return doc.GetElement(id); } catch { return null; } })
                    .Where(e => e != null && e.IsValidObject)
                    .ToList();

                if (validElements.Count == 0)
                {
                    TaskDialog.Show("Filter", "No valid elements found. Try refreshing the filter.");
                    SafeLoadTree();
                    return;
                }

                IEnumerable<Element> filtered = IsAndMode
                    ? validElements.Where(e => selectedItems.All(s => ElementMatchesCriteria(e, s, doc, projectUnits)))
                    : validElements.Where(e => selectedItems.Any(s => ElementMatchesCriteria(e, s, doc, projectUnits)));

                var idsToSelect = filtered.Select(e => e.Id).ToList();

                if (idsToSelect.Count > 0)
                {
                    uiDoc.Selection.SetElementIds(idsToSelect);
                    TopHeader = $"Filter ({idsToSelect.Count} matched)";
                }
                else
                {
                    TaskDialog.Show("Filter", "No elements match the selected criteria.");
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidObjectException)
            {
                TaskDialog.Show("Filter", "Some elements are no longer valid. Refreshing.");
                SafeLoadTree();
            }
        }

        private bool ElementMatchesCriteria(Element elem, SelectedItem item, Document doc, Units projectUnits)
        {
            try
            {
                if (elem?.Category == null) return false;

                // Category-only match (whole category was checked)
                if (item.IsCategoryMatch)
                    return item.CategoryId != null && elem.Category.Id == item.CategoryId;

                // Check category
                if (item.CategoryId != null && elem.Category.Id != item.CategoryId)
                    return false;
                if (!string.IsNullOrEmpty(item.CategoryName) && elem.Category.Name != item.CategoryName)
                    return false;

                // Check parameter value
                var param = elem.LookupParameter(item.PropertyName);
                if (param == null) return false;

                var paramValue = GetParameterValueSafe(param, doc, projectUnits);

                if (string.IsNullOrEmpty(item.PropertyValue) && string.IsNullOrEmpty(paramValue))
                    return true;

                return string.Equals(paramValue, item.PropertyValue, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static ElementId? GetParentCategoryId(TreeNodeViewModel node)
        {
            var parent = node.Parent;
            while (parent != null && parent.ElementId == null)
                parent = parent.Parent;
            return parent?.ElementId;
        }

        private static string GetParentCategoryName(TreeNodeViewModel node)
        {
            var parent = node.Parent;
            while (parent?.Parent != null)
                parent = parent.Parent;
            return parent?.Name ?? "";
        }

        #endregion

        #region Inner Types

        private sealed class SelectedItem
        {
            public string CategoryName { get; init; } = "";
            public ElementId? CategoryId { get; init; }
            public string PropertyName { get; init; } = "";
            public string PropertyValue { get; init; } = "";
            /// <summary>True when a whole category node was checked (match by category only).</summary>
            public bool IsCategoryMatch { get; init; }
        }

        #endregion
    }
}
