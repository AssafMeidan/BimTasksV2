using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Models;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// ViewModel for the Uniformat Generator window.
    /// Loads Uniformat codes from CSV, provides search/filter, and writes selected values
    /// to Revit shared parameters via IRevitUniformatWriter.
    /// </summary>
    public sealed class UniformatWindowViewModel : BindableBase
    {
        #region Services

        private readonly IUniformatDataService _data;
        private readonly IRevitUniformatWriter _writer;
        private readonly IRevitContextService _ctx;

        #endregion

        #region UI State

        public ObservableCollection<UniformatTreeItem> Tree { get; } = new();
        public ObservableCollection<BuiltInCategory> TargetCategories { get; } = new();

        /// <summary>Options for Scope ComboBox binding.</summary>
        public ApplyScope[] ScopeOptions { get; } = { ApplyScope.Types, ApplyScope.Instances };

        /// <summary>Options for TargetMode ComboBox binding.</summary>
        public TargetMode[] TargetModeOptions { get; } = { TargetMode.Categories, TargetMode.Selection };

        private string? _selectedCode;
        public string? SelectedCode { get => _selectedCode; set => SetProperty(ref _selectedCode, value); }

        private string? _selectedName;
        public string? SelectedName { get => _selectedName; set => SetProperty(ref _selectedName, value); }

        private string? _searchText;
        public string? SearchText
        {
            get => _searchText;
            set { if (SetProperty(ref _searchText, value)) RebuildTree(); }
        }

        private string? _rootFilter = "All";
        public string? RootFilter
        {
            get => _rootFilter;
            set { if (SetProperty(ref _rootFilter, value)) RebuildTree(); }
        }

        private int? _levelFilter;
        public int? LevelFilter
        {
            get => _levelFilter;
            set { if (SetProperty(ref _levelFilter, value)) RebuildTree(); }
        }

        private ApplyScope _scope = ApplyScope.Types;
        public ApplyScope Scope
        {
            get => _scope;
            set => SetProperty(ref _scope, value);
        }

        private TargetMode _targetMode = TargetMode.Categories;
        public TargetMode TargetMode
        {
            get => _targetMode;
            set => SetProperty(ref _targetMode, value);
        }

        private string _csvPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "BimTasks", "Uniformat_Clean.csv");
        public string CsvPath
        {
            get => _csvPath;
            set => SetProperty(ref _csvPath, value);
        }

        #endregion

        #region Commands

        public ICommand LoadCommand { get; }
        public ICommand EnsureParamsCommand { get; }
        public ICommand ApplyCommand { get; }
        public ICommand RefreshFilterCommand { get; }

        #endregion

        #region Backing Store

        private UniformatNode? _rootNode;

        #endregion

        #region Constructor

        public UniformatWindowViewModel()
        {
            var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
            _data = container.Resolve<IUniformatDataService>();
            _writer = container.Resolve<IRevitUniformatWriter>();
            _ctx = container.Resolve<IRevitContextService>();

            LoadCommand = new DelegateCommand(async () => await LoadAsync());
            EnsureParamsCommand = new DelegateCommand(EnsureParams);
            ApplyCommand = new DelegateCommand(Apply);
            RefreshFilterCommand = new DelegateCommand(RebuildTree);

            // Default target categories
            TargetCategories.Add(BuiltInCategory.OST_Walls);
            TargetCategories.Add(BuiltInCategory.OST_Doors);
        }

        #endregion

        #region Data Loading

        private async Task LoadAsync()
        {
            if (string.IsNullOrWhiteSpace(CsvPath))
            {
                TaskDialog.Show("BimTasks", "CSV path is empty.");
                return;
            }

            try
            {
                var flat = await _data.LoadFromCsvAsync(CsvPath).ConfigureAwait(false);
                _rootNode = _data.BuildTree(flat);
                RebuildTree();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[Uniformat] Failed to load CSV");
                TaskDialog.Show("BimTasks", $"Failed to load Uniformat CSV:\n{ex.Message}");
            }
        }

        private void RebuildTree()
        {
            Tree.Clear();
            if (_rootNode == null) return;

            foreach (var child in _rootNode.Children)
            {
                var filtered = UniformatTreeItem.FromNodeFiltered(
                    child,
                    OnSelect,
                    MatchesFilters);

                if (filtered != null)
                    Tree.Add(filtered);
            }
        }

        #endregion

        #region Filters

        private bool MatchesFilters(string code, string name, string level)
        {
            // Root filter
            if (!string.IsNullOrWhiteSpace(RootFilter) &&
                !RootFilter.Equals("All", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(code) || !code.StartsWith(RootFilter!, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            // Level filter
            if (LevelFilter.HasValue)
            {
                if (!int.TryParse(level, NumberStyles.Integer, CultureInfo.InvariantCulture, out int lv) ||
                    lv != LevelFilter.Value)
                    return false;
            }

            // Search text in code or name
            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                var txt = SearchText!.Trim();
                bool match =
                    (code?.IndexOf(txt, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0 ||
                    (name?.IndexOf(txt, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0;
                if (!match) return false;
            }

            return true;
        }

        #endregion

        #region Selection and Actions

        private void OnSelect(string code, string name)
        {
            SelectedCode = code;
            SelectedName = name;
        }

        private void EnsureParams()
        {
            var doc = _ctx.UIDocument?.Document;
            if (doc == null)
            {
                TaskDialog.Show("BimTasks", "No active document.");
                return;
            }

            try
            {
                _writer.EnsureParameters(doc, TargetCategories.ToList(), bindToTypes: Scope == ApplyScope.Types);
                TaskDialog.Show("BimTasks", "Shared parameters are ready.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[Uniformat] EnsureParameters failed");
                TaskDialog.Show("BimTasks", $"Failed to ensure parameters:\n{ex.Message}");
            }
        }

        private void Apply()
        {
            if (string.IsNullOrWhiteSpace(SelectedCode))
            {
                TaskDialog.Show("BimTasks", "Select a Uniformat item first.");
                return;
            }

            var uidoc = _ctx.UIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("BimTasks", "No active document.");
                return;
            }

            try
            {
                int updated = _writer.Apply(
                    uidoc,
                    SelectedCode!,
                    SelectedName ?? string.Empty,
                    Scope,
                    TargetMode,
                    TargetCategories.ToList());

                TaskDialog.Show("BimTasks", $"Updated {updated} elements.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[Uniformat] Apply failed");
                TaskDialog.Show("BimTasks", $"Apply failed:\n{ex.Message}");
            }
        }

        #endregion
    }

    /// <summary>
    /// UI tree item for Uniformat hierarchy with recursive builders and Select command.
    /// </summary>
    public sealed class UniformatTreeItem : BindableBase
    {
        public string Code { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Level { get; init; } = string.Empty;
        public ObservableCollection<UniformatTreeItem> Children { get; } = new();

        private readonly Action<string, string>? _onSelect;
        public ICommand SelectCommand { get; }

        private UniformatTreeItem(Action<string, string>? onSelect)
        {
            _onSelect = onSelect;
            SelectCommand = new DelegateCommand(() => _onSelect?.Invoke(Code, Name));
        }

        /// <summary>
        /// Build a full UI node (no filtering).
        /// </summary>
        public static UniformatTreeItem FromNode(UniformatNode node, Action<string, string> onSelect)
        {
            var ti = new UniformatTreeItem(onSelect)
            {
                Code = node.Code ?? string.Empty,
                Name = node.Name ?? string.Empty,
                Level = node.Level ?? string.Empty
            };
            foreach (var c in node.Children)
                ti.Children.Add(FromNode(c, onSelect));
            return ti;
        }

        /// <summary>
        /// Build a filtered UI node; returns null if this subtree does not match.
        /// </summary>
        public static UniformatTreeItem? FromNodeFiltered(
            UniformatNode node,
            Action<string, string> onSelect,
            Func<string, string, string, bool> predicate)
        {
            var childItems = new ObservableCollection<UniformatTreeItem>();
            foreach (var c in node.Children)
            {
                var child = FromNodeFiltered(c, onSelect, predicate);
                if (child != null) childItems.Add(child);
            }

            bool selfMatch = predicate(
                node.Code ?? string.Empty,
                node.Name ?? string.Empty,
                node.Level ?? string.Empty);

            if (!selfMatch && childItems.Count == 0)
                return null;

            var ti = new UniformatTreeItem(onSelect)
            {
                Code = node.Code ?? string.Empty,
                Name = node.Name ?? string.Empty,
                Level = node.Level ?? string.Empty
            };

            foreach (var ci in childItems)
                ti.Children.Add(ci);

            return ti;
        }
    }
}
