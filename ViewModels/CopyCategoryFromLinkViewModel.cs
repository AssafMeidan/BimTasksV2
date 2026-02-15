using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Prism.Commands;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// ViewModel for the CopyCategoryFromLink dialog.
    /// Provides cascading selections (Link -> Category -> Family -> Type)
    /// and computes element counts for copying from linked models.
    /// </summary>
    public class CopyCategoryFromLinkViewModel : BindableBase, IDisposable
    {
        private const string ServiceName = nameof(CopyCategoryFromLinkViewModel);
        private Document? _doc;
        private bool _disposed;

        public const string AllOption = "<All>";

        #region Properties - Links

        private ObservableCollection<RevitLinkInstance> _links = new();
        public ObservableCollection<RevitLinkInstance> Links
        {
            get => _links;
            set => SetProperty(ref _links, value);
        }

        private RevitLinkInstance? _selectedLink;
        public RevitLinkInstance? SelectedLink
        {
            get => _selectedLink;
            set
            {
                if (SetProperty(ref _selectedLink, value))
                {
                    Log.Debug("[{Service}] Selected link: {Link}", ServiceName, value?.Name ?? "null");
                    LoadCategoriesFromLink();
                    SelectedCategory = null;
                    UpdateElementCount();
                }
            }
        }

        #endregion

        #region Properties - Categories

        private ObservableCollection<CategoryItem> _categories = new();
        public ObservableCollection<CategoryItem> Categories
        {
            get => _categories;
            set => SetProperty(ref _categories, value);
        }

        private CategoryItem? _selectedCategory;
        public CategoryItem? SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                if (SetProperty(ref _selectedCategory, value))
                {
                    Log.Debug("[{Service}] Selected category: {Cat}", ServiceName, value?.Name ?? "null");
                    LoadFamiliesFromCategory();
                    UpdateElementCount();
                }
            }
        }

        #endregion

        #region Properties - Families

        private ObservableCollection<FamilyItem> _familyItems = new();
        public ObservableCollection<FamilyItem> FamilyItems
        {
            get => _familyItems;
            set => SetProperty(ref _familyItems, value);
        }

        private FamilyItem? _selectedFamily;
        public FamilyItem? SelectedFamily
        {
            get => _selectedFamily;
            set
            {
                if (SetProperty(ref _selectedFamily, value))
                {
                    Log.Debug("[{Service}] Selected family: {Fam}", ServiceName, value?.Name ?? "null");
                    LoadTypesFromFamily();
                    UpdateElementCount();
                }
            }
        }

        #endregion

        #region Properties - Types

        private ObservableCollection<CheckableItem> _typeItems = new();
        public ObservableCollection<CheckableItem> TypeItems
        {
            get => _typeItems;
            set => SetProperty(ref _typeItems, value);
        }

        #endregion

        #region Properties - UI State

        private int _totalElementCount;
        public int TotalElementCount
        {
            get => _totalElementCount;
            set => SetProperty(ref _totalElementCount, value);
        }

        private int _selectedElementCount;
        public int SelectedElementCount
        {
            get => _selectedElementCount;
            set => SetProperty(ref _selectedElementCount, value);
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        private string _statusMessage = "";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        #endregion

        #region Commands

        public DelegateCommand CopyCommand { get; }
        public DelegateCommand CancelCommand { get; }
        public DelegateCommand SelectAllTypesCommand { get; }
        public DelegateCommand SelectNoneTypesCommand { get; }
        public DelegateCommand RefreshCommand { get; }

        public Action? CloseAction { get; set; }
        public Action? CancelAction { get; set; }

        #endregion

        #region Constructor

        public CopyCategoryFromLinkViewModel()
        {
            CopyCommand = new DelegateCommand(OnCopy, CanCopy)
                .ObservesProperty(() => SelectedLink)
                .ObservesProperty(() => SelectedCategory);

            CancelCommand = new DelegateCommand(OnCancel);
            SelectAllTypesCommand = new DelegateCommand(OnSelectAllTypes, () => TypeItems.Any());
            SelectNoneTypesCommand = new DelegateCommand(OnSelectNoneTypes, () => TypeItems.Any());
            RefreshCommand = new DelegateCommand(OnRefresh);

            Log.Debug("[{Service}] Initialized", ServiceName);
        }

        #endregion

        #region Initialization

        public void Initialize(Document doc)
        {
            Log.Information("[{Service}] Initializing with document: {Doc}", ServiceName, doc?.Title ?? "null");

            // Clear all stale Revit object references
            _selectedLink = null;
            _selectedCategory = null;
            _selectedFamily = null;

            Links.Clear();
            Categories.Clear();
            FamilyItems.Clear();
            TypeItems.Clear();

            TotalElementCount = 0;
            SelectedElementCount = 0;
            StatusMessage = string.Empty;

            RaisePropertyChanged(nameof(SelectedLink));
            RaisePropertyChanged(nameof(SelectedCategory));
            RaisePropertyChanged(nameof(SelectedFamily));

            _doc = doc;
            LoadLinks();
        }

        #endregion

        #region Data Loading

        private void LoadLinks()
        {
            if (_doc == null) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading links...";

                var collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .Where(link => link.GetLinkDocument() != null)
                    .ToList();

                Links = new ObservableCollection<RevitLinkInstance>(collector);

                Log.Information("[{Service}] Found {Count} loaded links", ServiceName, Links.Count);

                if (Links.Any())
                {
                    SelectedLink = Links.First();
                }
                else
                {
                    StatusMessage = "No loaded links found in document";
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error loading links", ServiceName);
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void LoadCategoriesFromLink()
        {
            Categories.Clear();
            FamilyItems.Clear();
            TypeItems.Clear();

            if (SelectedLink == null) return;

            Document? linkDoc;
            try
            {
                linkDoc = SelectedLink.GetLinkDocument();
            }
            catch (Exception ex)
            {
                Log.Warning("[{Service}] Could not access link document: {Error}", ServiceName, ex.Message);
                _selectedLink = null;
                RaisePropertyChanged(nameof(SelectedLink));
                return;
            }

            if (linkDoc == null) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading categories...";

                var elementsInLink = new FilteredElementCollector(linkDoc)
                    .WhereElementIsNotElementType()
                    .WhereElementIsViewIndependent()
                    .ToElements();

                var categoryGroups = elementsInLink
                    .Where(e => e.Category != null && e.Category.CategoryType == CategoryType.Model)
                    .GroupBy(e => e.Category.Id)
                    .Select(g => new CategoryItem
                    {
                        Category = g.First().Category,
                        ElementCount = g.Count()
                    })
                    .OrderBy(c => c.Name)
                    .ToList();

                Categories = new ObservableCollection<CategoryItem>(categoryGroups);

                Log.Information("[{Service}] Found {Count} categories in link", ServiceName, Categories.Count);
                StatusMessage = $"Found {Categories.Count} categories";
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error loading categories", ServiceName);
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void LoadFamiliesFromCategory()
        {
            FamilyItems.Clear();
            TypeItems.Clear();

            if (SelectedLink == null || SelectedCategory == null) return;

            Document? linkDoc;
            try { linkDoc = SelectedLink.GetLinkDocument(); }
            catch { return; }
            if (linkDoc == null) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading families...";

                var elements = new FilteredElementCollector(linkDoc)
                    .OfCategoryId(SelectedCategory.Category!.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();

                var familyGroups = new Dictionary<string, int>();

                foreach (var elem in elements)
                {
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        var type = linkDoc.GetElement(typeId) as ElementType;
                        if (type != null && !string.IsNullOrEmpty(type.FamilyName))
                        {
                            if (!familyGroups.ContainsKey(type.FamilyName))
                                familyGroups[type.FamilyName] = 0;
                            familyGroups[type.FamilyName]++;
                        }
                    }
                }

                var familyItems = familyGroups
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => new FamilyItem { Name = kvp.Key, ElementCount = kvp.Value })
                    .ToList();

                int totalCount = familyItems.Sum(f => f.ElementCount);
                familyItems.Insert(0, new FamilyItem { Name = AllOption, ElementCount = totalCount });

                FamilyItems = new ObservableCollection<FamilyItem>(familyItems);
                SelectedFamily = FamilyItems.First();

                Log.Information("[{Service}] Found {Count} families", ServiceName, FamilyItems.Count - 1);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error loading families", ServiceName);
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void LoadTypesFromFamily()
        {
            TypeItems.Clear();
            SelectAllTypesCommand.RaiseCanExecuteChanged();
            SelectNoneTypesCommand.RaiseCanExecuteChanged();

            if (SelectedLink == null || SelectedCategory == null || SelectedFamily == null)
                return;

            if (SelectedFamily.Name == AllOption)
            {
                StatusMessage = "All families selected - all types will be copied";
                return;
            }

            Document? linkDoc;
            try { linkDoc = SelectedLink.GetLinkDocument(); }
            catch { return; }
            if (linkDoc == null) return;

            try
            {
                IsLoading = true;
                StatusMessage = "Loading types...";

                var elements = new FilteredElementCollector(linkDoc)
                    .OfCategoryId(SelectedCategory.Category!.Id)
                    .WhereElementIsNotElementType()
                    .ToElements();

                var typeGroups = new Dictionary<string, int>();

                foreach (var elem in elements)
                {
                    ElementId typeId = elem.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                    {
                        var type = linkDoc.GetElement(typeId) as ElementType;
                        if (type != null && type.FamilyName == SelectedFamily.Name)
                        {
                            if (!string.IsNullOrEmpty(type.Name))
                            {
                                if (!typeGroups.ContainsKey(type.Name))
                                    typeGroups[type.Name] = 0;
                                typeGroups[type.Name]++;
                            }
                        }
                    }
                }

                var checkables = typeGroups
                    .OrderBy(kvp => kvp.Key)
                    .Select(kvp => new CheckableItem
                    {
                        Name = kvp.Key,
                        IsChecked = true,
                        ElementCount = kvp.Value
                    })
                    .ToList();

                foreach (var item in checkables)
                {
                    item.PropertyChanged += (s, e) =>
                    {
                        if (e.PropertyName == nameof(CheckableItem.IsChecked))
                            UpdateElementCount();
                    };
                }

                TypeItems = new ObservableCollection<CheckableItem>(checkables);

                SelectAllTypesCommand.RaiseCanExecuteChanged();
                SelectNoneTypesCommand.RaiseCanExecuteChanged();

                Log.Information("[{Service}] Found {Count} types", ServiceName, TypeItems.Count);
                StatusMessage = $"Found {TypeItems.Count} types";
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error loading types", ServiceName);
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        #endregion

        #region Element Count

        private void UpdateElementCount()
        {
            if (SelectedLink == null || SelectedCategory == null)
            {
                TotalElementCount = 0;
                SelectedElementCount = 0;
                return;
            }

            if (SelectedFamily == null || SelectedFamily.Name == AllOption)
            {
                TotalElementCount = SelectedCategory.ElementCount;
                SelectedElementCount = TotalElementCount;
            }
            else
            {
                TotalElementCount = SelectedFamily.ElementCount;

                if (TypeItems.Any())
                {
                    SelectedElementCount = TypeItems
                        .Where(t => t.IsChecked)
                        .Sum(t => t.ElementCount);
                }
                else
                {
                    SelectedElementCount = TotalElementCount;
                }
            }

            StatusMessage = $"Will copy {SelectedElementCount} of {TotalElementCount} elements";
        }

        #endregion

        #region Command Handlers

        private bool CanCopy()
        {
            return SelectedLink != null && SelectedCategory != null;
        }

        private void OnCopy()
        {
            Log.Information("[{Service}] Copy requested - {Count} elements", ServiceName, SelectedElementCount);
            CloseAction?.Invoke();
        }

        private void OnCancel()
        {
            Log.Debug("[{Service}] Cancel requested", ServiceName);
            CancelAction?.Invoke();
        }

        private void OnSelectAllTypes()
        {
            foreach (var item in TypeItems)
                item.IsChecked = true;
        }

        private void OnSelectNoneTypes()
        {
            foreach (var item in TypeItems)
                item.IsChecked = false;
        }

        private void OnRefresh()
        {
            if (_doc != null)
                Initialize(_doc);
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets the element IDs to copy based on current selection.
        /// </summary>
        public List<ElementId> GetElementIdsToCopy()
        {
            if (SelectedLink == null || SelectedCategory == null)
                return new List<ElementId>();

            Document? linkDoc;
            try { linkDoc = SelectedLink.GetLinkDocument(); }
            catch { return new List<ElementId>(); }
            if (linkDoc == null) return new List<ElementId>();

            var elements = new FilteredElementCollector(linkDoc)
                .OfCategoryId(SelectedCategory.Category!.Id)
                .WhereElementIsNotElementType()
                .ToElements();

            bool copyAllFamilies = SelectedFamily == null || SelectedFamily.Name == AllOption;

            if (copyAllFamilies)
                return elements.Select(e => e.Id).ToList();

            var checkedTypes = new HashSet<string>(
                TypeItems.Where(t => t.IsChecked).Select(t => t.Name));

            var result = new List<ElementId>();

            foreach (var elem in elements)
            {
                ElementId typeId = elem.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    var type = linkDoc.GetElement(typeId) as ElementType;
                    if (type != null && type.FamilyName == SelectedFamily!.Name)
                    {
                        if (checkedTypes.Contains(type.Name))
                            result.Add(elem.Id);
                    }
                }
            }

            return result;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                _selectedLink = null;
                _selectedCategory = null;
                _selectedFamily = null;
                _doc = null;

                Links?.Clear();
                Categories?.Clear();
                FamilyItems?.Clear();
                TypeItems?.Clear();

                Log.Debug("[{Service}] Disposed", ServiceName);
            }

            _disposed = true;
        }

        #endregion
    }

    #region Helper Classes

    /// <summary>
    /// Category with element count for display in ComboBox.
    /// </summary>
    public class CategoryItem : BindableBase
    {
        public Category? Category { get; set; }
        public string Name => Category?.Name ?? string.Empty;
        public int ElementCount { get; set; }
        public string DisplayName => $"{Name} ({ElementCount})";
    }

    /// <summary>
    /// Family with element count for display in ComboBox.
    /// </summary>
    public class FamilyItem : BindableBase
    {
        public string Name { get; set; } = "";
        public int ElementCount { get; set; }
        public string DisplayName => $"{Name} ({ElementCount})";
    }

    /// <summary>
    /// Checkable item for type selection with element count.
    /// </summary>
    public class CheckableItem : BindableBase
    {
        private bool _isChecked;
        public bool IsChecked
        {
            get => _isChecked;
            set => SetProperty(ref _isChecked, value);
        }

        private string _name = "";
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        public int ElementCount { get; set; }
        public string DisplayName => $"{Name} ({ElementCount})";
    }

    #endregion
}
