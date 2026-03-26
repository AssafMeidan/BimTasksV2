using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Autodesk.Revit.DB;
using BimTasksV2.Helpers;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class CategoryCheckItem : BindableBase
    {
        public BuiltInCategory Category { get; set; }

        private string _name = "";
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        private int _count;
        public int Count
        {
            get => _count;
            set => SetProperty(ref _count, value);
        }

        private bool _isChecked;
        public bool IsChecked
        {
            get => _isChecked;
            set => SetProperty(ref _isChecked, value);
        }
    }

    public class SwatchRowItem : BindableBase
    {
        public string Value { get; set; } = "";
        public string Description { get; set; } = "";
        public int Count { get; set; }
        public List<ElementId> ElementIds { get; set; } = new();

        private System.Windows.Media.Color _color;
        public System.Windows.Media.Color Color
        {
            get => _color;
            set
            {
                SetProperty(ref _color, value);
                RaisePropertyChanged(nameof(ColorBrush));
            }
        }

        public SolidColorBrush ColorBrush => new(Color);
    }

    public class ColorSwatchViewModel : BindableBase
    {
        private readonly ColorOverrideService _overrideService = new();
        private Autodesk.Revit.UI.UIApplication? _uiApp;
        private View? _activeView;

        // Lazy-resolved: registered in OnFirstIdle()
        private ICommandDispatcherService? _dispatcher;
        private ICommandDispatcherService Dispatcher =>
            _dispatcher ??= BimTasksV2.Infrastructure.ContainerLocator.Container
                .Resolve<ICommandDispatcherService>();

        // 16-color distinguishable categorical palette
        private static readonly System.Windows.Media.Color[] DefaultPalette = new[]
        {
            System.Windows.Media.Color.FromRgb(0x1F, 0x77, 0xB4), // blue
            System.Windows.Media.Color.FromRgb(0xFF, 0x7F, 0x0E), // orange
            System.Windows.Media.Color.FromRgb(0x2C, 0xA0, 0x2C), // green
            System.Windows.Media.Color.FromRgb(0xD6, 0x27, 0x28), // red
            System.Windows.Media.Color.FromRgb(0x94, 0x67, 0xBD), // purple
            System.Windows.Media.Color.FromRgb(0x8C, 0x56, 0x4B), // brown
            System.Windows.Media.Color.FromRgb(0xE3, 0x77, 0xC2), // pink
            System.Windows.Media.Color.FromRgb(0x17, 0xBE, 0xCF), // teal
            System.Windows.Media.Color.FromRgb(0xBC, 0xBD, 0x22), // olive
            System.Windows.Media.Color.FromRgb(0xAE, 0xC7, 0xE8), // light blue
            System.Windows.Media.Color.FromRgb(0xFF, 0xBB, 0x78), // light orange
            System.Windows.Media.Color.FromRgb(0x98, 0xDF, 0x8A), // light green
            System.Windows.Media.Color.FromRgb(0xFF, 0x98, 0x96), // light red
            System.Windows.Media.Color.FromRgb(0xC5, 0xB0, 0xD5), // light purple
            System.Windows.Media.Color.FromRgb(0xC4, 0x9C, 0x94), // light brown
            System.Windows.Media.Color.FromRgb(0x9E, 0xDA, 0xE5), // light teal
        };

        // Preset palette for color picker popup
        public static readonly System.Windows.Media.Color[] PresetColors = new[]
        {
            System.Windows.Media.Color.FromRgb(0xD3, 0x2F, 0x2F), // Red
            System.Windows.Media.Color.FromRgb(0x19, 0x76, 0xD2), // Blue
            System.Windows.Media.Color.FromRgb(0x38, 0x8E, 0x3C), // Green
            System.Windows.Media.Color.FromRgb(0xF5, 0x7C, 0x00), // Orange
            System.Windows.Media.Color.FromRgb(0x7B, 0x1F, 0xA2), // Purple
            System.Windows.Media.Color.FromRgb(0x00, 0x83, 0x8F), // Teal
            System.Windows.Media.Color.FromRgb(0xC2, 0x18, 0x5B), // Pink
            System.Windows.Media.Color.FromRgb(0x5D, 0x40, 0x37), // Brown
            System.Windows.Media.Color.FromRgb(0xFF, 0xB3, 0x00), // Amber
            System.Windows.Media.Color.FromRgb(0x45, 0x5A, 0x64), // Blue Grey
        };

        public ObservableCollection<string> ParameterNames { get; } = new();
        public ObservableCollection<string> DescriptionParameterNames { get; } = new();
        public ObservableCollection<CategoryCheckItem> Categories { get; } = new();
        public ObservableCollection<SwatchRowItem> SwatchRows { get; } = new();

        private string? _selectedParameter;
        public string? SelectedParameter
        {
            get => _selectedParameter;
            set
            {
                if (SetProperty(ref _selectedParameter, value) && value != null)
                    OnParameterChanged();
            }
        }

        private string? _selectedDescriptionParameter;
        public string? SelectedDescriptionParameter
        {
            get => _selectedDescriptionParameter;
            set
            {
                if (SetProperty(ref _selectedDescriptionParameter, value))
                    OnParameterChanged();
            }
        }

        private bool _isCategoryDropdownOpen;
        public bool IsCategoryDropdownOpen
        {
            get => _isCategoryDropdownOpen;
            set => SetProperty(ref _isCategoryDropdownOpen, value);
        }

        public string CategorySummary
        {
            get
            {
                var checked_ = Categories.Where(c => c.IsChecked).ToList();
                if (checked_.Count == 0) return "None selected";
                if (checked_.Count <= 2) return string.Join(", ", checked_.Select(c => c.Name));
                return $"{checked_.Count} categories";
            }
        }

        // Preserved color map for refresh — value → color
        private Dictionary<string, System.Windows.Media.Color>? _preservedColors;

        public DelegateCommand ApplyColorsCommand { get; }
        public DelegateCommand ClearColorsCommand { get; }
        public DelegateCommand RefreshCommand { get; }
        public DelegateCommand ResetViewCommand { get; }
        public DelegateCommand CreateLegendCommand { get; }
        public DelegateCommand ToggleCategoryDropdownCommand { get; }
        public DelegateCommand<SwatchRowItem> SelectElementsCommand { get; }

        public ColorSwatchViewModel()
        {
            ApplyColorsCommand = new DelegateCommand(OnApplyColors, CanApplyColors);
            ClearColorsCommand = new DelegateCommand(OnClearColors, CanClearColors);
            RefreshCommand = new DelegateCommand(OnRefresh, () => _selectedParameter != null);
            ResetViewCommand = new DelegateCommand(OnResetView);
            CreateLegendCommand = new DelegateCommand(OnCreateLegend, () => SwatchRows.Count > 0);
            ToggleCategoryDropdownCommand = new DelegateCommand(() =>
                IsCategoryDropdownOpen = !IsCategoryDropdownOpen);
            SelectElementsCommand = new DelegateCommand<SwatchRowItem>(OnSelectElements);
        }

        /// <summary>
        /// Called by the handler after switching to this panel view.
        /// Sets up the UI with data from the active Revit view.
        /// </summary>
        public void Initialize(Autodesk.Revit.UI.UIApplication uiApp)
        {
            _uiApp = uiApp;
            var doc = uiApp.ActiveUIDocument?.Document;
            var view = doc?.ActiveView;

            if (doc == null || view == null)
            {
                Log.Warning("[ColorSwatch] No active document or view");
                return;
            }

            _activeView = view;

            // Load categories
            Categories.Clear();
            var cats = _overrideService.GetModelCategories(doc, view);
            int totalElements = cats.Sum(c => c.Count);
            int threshold = Math.Max(1, (int)(totalElements * 0.02)); // 2% threshold

            foreach (var (cat, name, count) in cats)
            {
                var item = new CategoryCheckItem
                {
                    Category = cat,
                    Name = name,
                    Count = count,
                    IsChecked = count >= threshold, // auto-check categories with 2%+ of elements
                };
                item.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(CategoryCheckItem.IsChecked))
                    {
                        RaisePropertyChanged(nameof(CategorySummary));
                        RefreshParameterList();
                    }
                };
                Categories.Add(item);
            }

            RaisePropertyChanged(nameof(CategorySummary));
            RefreshParameterList();
        }

        private void RefreshParameterList()
        {
            if (_activeView == null) return;

            var selectedCats = GetSelectedCategories();
            if (selectedCats.Count == 0)
            {
                ParameterNames.Clear();
                DescriptionParameterNames.Clear();
                return;
            }

            var viewId = _activeView.Id;
            var previousParam = _selectedParameter;
            var previousDesc = _selectedDescriptionParameter;

            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    if (doc == null) return;
                    var view = doc.GetElement(viewId) as View;
                    if (view == null) return;

                    var names = _overrideService.GetSharedParameterNames(doc, view, selectedCats);

                    // Find a default description parameter (first one with a value)
                    string? defaultDesc = null;
                    if (previousDesc == null || previousDesc == "None")
                    {
                        defaultDesc = _overrideService.FindFirstParameterWithValue(
                            doc, view, selectedCats, names, previousParam);
                    }

                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        ParameterNames.Clear();
                        foreach (var name in names)
                            ParameterNames.Add(name);

                        DescriptionParameterNames.Clear();
                        DescriptionParameterNames.Add("None");
                        foreach (var name in names)
                            DescriptionParameterNames.Add(name);

                        if (previousParam != null && ParameterNames.Contains(previousParam))
                            SelectedParameter = previousParam;

                        if (previousDesc != null && previousDesc != "None"
                            && DescriptionParameterNames.Contains(previousDesc))
                            SelectedDescriptionParameter = previousDesc;
                        else if (defaultDesc != null)
                            SelectedDescriptionParameter = defaultDesc;
                        else
                            SelectedDescriptionParameter = "None";
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to refresh parameter list");
                }
            });
        }

        private void OnParameterChanged()
        {
            if (_activeView == null || _selectedParameter == null) return;

            var selectedCats = GetSelectedCategories();
            if (selectedCats.Count == 0) return;

            var viewId = _activeView.Id;
            var paramName = _selectedParameter;
            string? descParam = _selectedDescriptionParameter == "None"
                ? null
                : _selectedDescriptionParameter;

            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    if (doc == null) return;
                    var view = doc.GetElement(viewId) as View;
                    if (view == null) return;

                    var groups = _overrideService.ScanElements(
                        doc, view, paramName, descParam, selectedCats);

                    // Capture preserved colors before dispatching to UI thread
                    var colorMap = _preservedColors;
                    _preservedColors = null;

                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        SwatchRows.Clear();
                        for (int i = 0; i < groups.Count; i++)
                        {
                            var g = groups[i];
                            System.Windows.Media.Color color;
                            if (colorMap != null && colorMap.TryGetValue(g.Value, out var preserved))
                                color = preserved;
                            else if (g.Value == "(empty)")
                                color = System.Windows.Media.Color.FromRgb(0xBD, 0xBD, 0xBD);
                            else
                                color = DefaultPalette[i % DefaultPalette.Length];

                            SwatchRows.Add(new SwatchRowItem
                            {
                                Value = g.Value,
                                Description = g.Description,
                                Count = g.ElementIds.Count,
                                ElementIds = g.ElementIds,
                                Color = color,
                            });
                        }

                        ApplyColorsCommand.RaiseCanExecuteChanged();
                        RefreshCommand.RaiseCanExecuteChanged();
                        CreateLegendCommand.RaiseCanExecuteChanged();
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to scan elements");
                }
            });
        }

        private bool CanApplyColors() => SwatchRows.Count > 0;
        private bool CanClearColors() => _overrideService.HasActiveOverrides;

        private void OnApplyColors()
        {
            // Build groups snapshot on UI thread
            var groups = SwatchRows.Select(r => new ColorSwatchGroup
            {
                Value = r.Value,
                Description = r.Description,
                Color = r.Color,
                ElementIds = r.ElementIds,
            }).ToList();

            // Execute on Revit API thread
            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    var view = doc?.ActiveView;
                    if (doc == null || view == null) return;

                    _overrideService.ApplyOverrides(doc, view, groups);

                    // Update UI on dispatcher
                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                        ClearColorsCommand.RaiseCanExecuteChanged());
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to apply colors");
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasksV2", $"Failed to apply colors:\n{ex.Message}");
                }
            });
        }

        private void OnClearColors()
        {
            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    var view = doc?.ActiveView;
                    if (doc == null || view == null) return;

                    _overrideService.ClearOverrides(doc, view);

                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                        ClearColorsCommand.RaiseCanExecuteChanged());
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to clear colors");
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasksV2", $"Failed to clear colors:\n{ex.Message}");
                }
            });
        }

        private void OnRefresh()
        {
            // Preserve current color assignments so the refresh keeps them
            _preservedColors = SwatchRows.ToDictionary(r => r.Value, r => r.Color);
            OnParameterChanged();
        }

        private void OnCreateLegend()
        {
            var paramName = _selectedParameter ?? "Swatch";

            // Build legend entries on UI thread
            var entries = SwatchRows.Select(r => new LegendEntry
            {
                Color = new Autodesk.Revit.DB.Color(r.Color.R, r.Color.G, r.Color.B),
                Label = r.Value,
                SecondaryLabel = string.IsNullOrEmpty(r.Description) ? null : r.Description,
            }).ToList();

            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    if (doc == null) return;

                    var builder = new LegendBuilder();
                    View? legendView;

                    using (var tg = new TransactionGroup(doc, "Color Swatch — Create Legend"))
                    {
                        tg.Start();
                        legendView = builder.CreateLegend(doc, $"Swatch - {paramName}", entries);
                        tg.Assimilate();
                    }

                    if (legendView != null)
                    {
                        // Navigate to the new legend
                        uiApp.ActiveUIDocument.ActiveView = legendView;
                        Log.Information("[ColorSwatch] Created legend '{Name}' with {Count} entries",
                            legendView.Name, entries.Count);
                    }
                    else
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("BimTasksV2",
                            "Failed to create legend view. Make sure a legend named 'empty' exists, or a Drafting view type is available.");
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to create legend");
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasksV2", $"Failed to create legend:\n{ex.Message}");
                }
            });
        }

        private void OnResetView()
        {
            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    var view = doc?.ActiveView;
                    if (doc == null || view == null) return;

                    int count = _overrideService.ClearAllSwatchOverrides(doc, view);

                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        ClearColorsCommand.RaiseCanExecuteChanged();
                        if (count > 0)
                            Log.Information("[ColorSwatch] Reset {Count} overrides in view '{View}'",
                                count, view.Name);
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to reset view overrides");
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasksV2",
                        $"Failed to reset view overrides:\n{ex.Message}");
                }
            });
        }

        private void OnSelectElements(SwatchRowItem? row)
        {
            if (row == null || row.ElementIds.Count == 0) return;

            var ids = row.ElementIds.ToList();
            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var uidoc = uiApp.ActiveUIDocument;
                    if (uidoc == null) return;

                    uidoc.Selection.SetElementIds(ids);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[ColorSwatch] Failed to select elements");
                }
            });
        }

        /// <summary>
        /// Called when the tab is closed — clears overrides.
        /// </summary>
        public void OnTabClosed()
        {
            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var doc = uiApp.ActiveUIDocument?.Document;
                    var view = doc?.ActiveView;
                    if (doc == null || view == null) return;

                    _overrideService.ClearOverrides(doc, view);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[ColorSwatch] Failed to clear overrides on tab close");
                }
            });
        }

        /// <summary>
        /// Updates a swatch row's color (called from color picker in view).
        /// </summary>
        public void UpdateSwatchColor(SwatchRowItem row, System.Windows.Media.Color newColor)
        {
            row.Color = newColor;
        }

        private List<BuiltInCategory> GetSelectedCategories()
        {
            return Categories
                .Where(c => c.IsChecked)
                .Select(c => c.Category)
                .ToList();
        }
    }
}
