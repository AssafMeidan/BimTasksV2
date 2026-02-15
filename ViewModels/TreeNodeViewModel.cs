using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Prism.Mvvm;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// Generic checkbox tree node used by FilterTreeView and other hierarchical views.
    /// Supports cascading check/uncheck to children.
    /// </summary>
    public class TreeNodeViewModel : BindableBase
    {
        private string _name = "";
        public string Name { get => _name; set => SetProperty(ref _name, value); }

        private string? _displayName;
        public string DisplayName { get => _displayName ?? _name; set => SetProperty(ref _displayName, value); }

        private bool _isChecked;
        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (SetProperty(ref _isChecked, value) && Children != null)
                {
                    foreach (var child in Children)
                        child.IsChecked = value;
                }
            }
        }

        public TreeNodeViewModel? Parent { get; set; }
        public ElementId? ElementId { get; set; }
        public string? PropertyName { get; set; }
        public string? PropertyValue { get; set; }
        public ObservableCollection<TreeNodeViewModel> Children { get; } = new();

        public TreeNodeViewModel() { }
        public TreeNodeViewModel(string name, TreeNodeViewModel? parent = null)
        {
            _name = name;
            Parent = parent;
        }
    }
}
