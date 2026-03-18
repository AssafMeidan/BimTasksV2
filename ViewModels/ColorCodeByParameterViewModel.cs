using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Autodesk.Revit.DB;
using BimTasksV2.Events;
using BimTasksV2.Helpers.ColorCodeByParameter;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class LegendItem : BindableBase
    {
        public string PrimaryValue { get; set; } = "";
        public string SecondaryValue { get; set; } = "";

        private System.Windows.Media.Color _swatchColor;
        public System.Windows.Media.Color SwatchColor
        {
            get => _swatchColor;
            set => SetProperty(ref _swatchColor, value);
        }

        public int Count { get; set; }
        public string CountText => $"({Count})";
        public List<ElementId> ElementIds { get; set; } = new();

        public System.Windows.Visibility SecondaryVisibility =>
            string.IsNullOrEmpty(SecondaryValue) ? System.Windows.Visibility.Collapsed : System.Windows.Visibility.Visible;
    }

    public class ColorCodeByParameterViewModel : BindableBase
    {
        private readonly ICommandDispatcherService _dispatcher;
        private List<ElementId> _trackedElementIds = new();
        private string? _lastPrimaryParam;
        private string? _lastSecondaryParam;

        public DelegateCommand ApplyCommand { get; }
        public DelegateCommand RefreshCommand { get; }
        public DelegateCommand ResetCommand { get; }
        public DelegateCommand ShuffleColorsCommand { get; }

        public ObservableCollection<string> ParameterNames { get; } = new();
        public ObservableCollection<LegendItem> LegendItems { get; } = new();
        public ObservableCollection<System.Windows.Media.Color> PaletteColors { get; } = new();

        public ColorCodeByParameterViewModel()
        {
            var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
            _dispatcher = container.Resolve<ICommandDispatcherService>();

            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.ColorCodeInitEvent>()
                .Subscribe(OnColorCodeInit, ThreadOption.UIThread);

            ApplyCommand = new DelegateCommand(ExecuteApply, () => !string.IsNullOrEmpty(SelectedPrimaryParam));
            RefreshCommand = new DelegateCommand(ExecuteRefresh, () => _lastPrimaryParam != null);
            ResetCommand = new DelegateCommand(ExecuteReset, () => _trackedElementIds.Count > 0);
            ShuffleColorsCommand = new DelegateCommand(ExecuteShuffle, () => LegendItems.Count > 0);

            // Populate palette colors for the color picker
            foreach (var c in ColorCodeService.GetPalette())
                PaletteColors.Add(System.Windows.Media.Color.FromRgb(c.Red, c.Green, c.Blue));
        }

        #region Properties

        private string? _selectedPrimaryParam;
        public string? SelectedPrimaryParam
        {
            get => _selectedPrimaryParam;
            set { SetProperty(ref _selectedPrimaryParam, value); ApplyCommand.RaiseCanExecuteChanged(); }
        }

        private string? _selectedSecondaryParam;
        public string? SelectedSecondaryParam
        {
            get => _selectedSecondaryParam;
            set => SetProperty(ref _selectedSecondaryParam, value);
        }

        private bool _useSelection = true;
        public bool UseSelection
        {
            get => _useSelection;
            set => SetProperty(ref _useSelection, value);
        }

        private bool _useAllInView;
        public bool UseAllInView
        {
            get => _useAllInView;
            set => SetProperty(ref _useAllInView, value);
        }

        private string _statusText = "Select parameters and click Apply.";
        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        private LegendItem? _editingLegendItem;
        public LegendItem? EditingLegendItem
        {
            get => _editingLegendItem;
            set => SetProperty(ref _editingLegendItem, value);
        }

        private LegendItem? _selectedLegendItem;
        public LegendItem? SelectedLegendItem
        {
            get => _selectedLegendItem;
            set
            {
                SetProperty(ref _selectedLegendItem, value);
                if (value != null)
                    SelectLegendElements(value);
            }
        }

        #endregion

        private void OnColorCodeInit(object? _)
        {
            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var uidoc = uiApp.ActiveUIDocument;
                    var doc = uidoc.Document;

                    // Get element IDs to scan for parameter names
                    var elementIds = GetTargetElementIds(uidoc);
                    if (elementIds.Count == 0)
                    {
                        // Fall back to all elements in view
                        elementIds = new FilteredElementCollector(doc, doc.ActiveView.Id)
                            .WhereElementIsNotElementType()
                            .ToElementIds()
                            .ToList();
                    }

                    var names = ColorCodeService.GetParameterNames(doc, elementIds);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ParameterNames.Clear();
                        foreach (var name in names)
                            ParameterNames.Add(name);

                        StatusText = $"Found {names.Count} parameters from {elementIds.Count} elements.";
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorCodeVM] Init failed");
                    Application.Current.Dispatcher.Invoke(() =>
                        StatusText = $"Error scanning parameters: {ex.Message}");
                }
            });
        }

        private void ExecuteApply()
        {
            var primary = SelectedPrimaryParam;
            if (string.IsNullOrEmpty(primary)) return;

            var secondary = SelectedSecondaryParam;
            _lastPrimaryParam = primary;
            _lastSecondaryParam = secondary;
            StatusText = "Applying colors...";

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var uidoc = uiApp.ActiveUIDocument;
                    var doc = uidoc.Document;
                    var view = doc.ActiveView;

                    var elementIds = GetTargetElementIds(uidoc);
                    if (elementIds.Count == 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                            StatusText = "No elements found. Select elements or switch to 'All in View'.");
                        return;
                    }

                    var results = ColorCodeService.ApplyColorOverrides(
                        doc, view, primary, secondary, elementIds);

                    // Track all element IDs for reset
                    _trackedElementIds = elementIds;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        LegendItems.Clear();
                        foreach (var r in results)
                        {
                            LegendItems.Add(new LegendItem
                            {
                                PrimaryValue = r.PrimaryValue,
                                SecondaryValue = r.SecondaryValue,
                                SwatchColor = System.Windows.Media.Color.FromRgb(r.Color.Red, r.Color.Green, r.Color.Blue),
                                Count = r.ElementIds.Count,
                                ElementIds = r.ElementIds
                            });
                        }

                        StatusText = $"Applied {results.Count} color groups to {elementIds.Count} elements.";
                        RefreshCommand.RaiseCanExecuteChanged();
                        ResetCommand.RaiseCanExecuteChanged();
                        ShuffleColorsCommand.RaiseCanExecuteChanged();
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorCodeVM] Apply failed");
                    Application.Current.Dispatcher.Invoke(() =>
                        StatusText = $"Error: {ex.Message}");
                }
            });
        }

        private void ExecuteRefresh()
        {
            // Re-apply with same parameters
            SelectedPrimaryParam = _lastPrimaryParam;
            SelectedSecondaryParam = _lastSecondaryParam;
            ExecuteApply();
        }

        private void ExecuteReset()
        {
            var idsToReset = new List<ElementId>(_trackedElementIds);
            StatusText = "Resetting colors...";

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    var view = doc.ActiveView;
                    ColorCodeService.ClearAllOverrides(doc, view, idsToReset);

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        LegendItems.Clear();
                        _trackedElementIds.Clear();
                        StatusText = "All color overrides cleared.";
                        ResetCommand.RaiseCanExecuteChanged();
                        ShuffleColorsCommand.RaiseCanExecuteChanged();
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorCodeVM] Reset failed");
                    Application.Current.Dispatcher.Invoke(() =>
                        StatusText = $"Error resetting: {ex.Message}");
                }
            });
        }

        private void ExecuteShuffle()
        {
            if (LegendItems.Count == 0) return;

            var shuffled = ColorCodeService.GetShuffledDistancedColors(LegendItems.Count);
            for (int i = 0; i < LegendItems.Count; i++)
            {
                var c = shuffled[i];
                LegendItems[i].SwatchColor = System.Windows.Media.Color.FromRgb(c.Red, c.Green, c.Blue);
            }

            ReapplyAllOverrides();
            StatusText = "Colors reshuffled.";
        }

        public void ApplyColorToItem(System.Windows.Media.Color newColor)
        {
            if (EditingLegendItem == null) return;

            EditingLegendItem.SwatchColor = newColor;
            var elementIds = new List<ElementId>(EditingLegendItem.ElementIds);
            var revitColor = new Autodesk.Revit.DB.Color(newColor.R, newColor.G, newColor.B);
            EditingLegendItem = null;

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    var view = doc.ActiveView;
                    ColorCodeService.ApplyGroupOverride(doc, view, elementIds, revitColor);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[ColorCodeVM] Color change failed");
                }
            });
        }

        private void ReapplyAllOverrides()
        {
            var groups = LegendItems.Select(li => new
            {
                ElementIds = new List<ElementId>(li.ElementIds),
                li.SwatchColor
            }).ToList();

            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument.Document;
                    var view = doc.ActiveView;
                    ColorCodeService.ApplyMultipleGroupOverrides(doc, view,
                        groups.Select(g => (
                            (ICollection<ElementId>)g.ElementIds,
                            new Autodesk.Revit.DB.Color(g.SwatchColor.R, g.SwatchColor.G, g.SwatchColor.B)
                        )).ToList());
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorCodeVM] Reapply failed");
                    Application.Current.Dispatcher.Invoke(() =>
                        StatusText = $"Error applying colors: {ex.Message}");
                }
            });
        }

        private void SelectLegendElements(LegendItem item)
        {
            _dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    ColorCodeService.SelectElements(uiApp.ActiveUIDocument, item.ElementIds);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[ColorCodeVM] Selection failed");
                }
            });
        }

        private List<ElementId> GetTargetElementIds(Autodesk.Revit.UI.UIDocument uidoc)
        {
            if (_useSelection)
            {
                var selected = uidoc.Selection.GetElementIds().ToList();
                if (selected.Count > 0) return selected;
            }

            // All elements in active view (excluding element types)
            return new FilteredElementCollector(uidoc.Document, uidoc.Document.ActiveView.Id)
                .WhereElementIsNotElementType()
                .ToElementIds()
                .ToList();
        }
    }
}
