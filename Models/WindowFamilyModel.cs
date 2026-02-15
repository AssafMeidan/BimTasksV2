using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Prism.Mvvm;

namespace BimTasksV2.Models
{
    /// <summary>
    /// Represents a window Family for the CreateWindowFamilies dialog.
    /// </summary>
    public class WindowFamilyModel : BindableBase
    {
        /// <summary>Display name of the family.</summary>
        public string Name { get; set; } = "";

        /// <summary>The Revit Family object.</summary>
        public Family? Family { get; set; }

        /// <summary>Symbol (type) definitions belonging to this family.</summary>
        public ObservableCollection<WindowSymbolModel> Symbols { get; } = new();
    }

    /// <summary>
    /// Represents a single window FamilySymbol (type) with its key dimensions.
    /// </summary>
    public class WindowSymbolModel : BindableBase
    {
        private string _name = "";
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        private double _width;
        public double Width
        {
            get => _width;
            set => SetProperty(ref _width, value);
        }

        private double _height;
        public double Height
        {
            get => _height;
            set => SetProperty(ref _height, value);
        }

        private double _defaultSillHeight;
        public double DefaultSillHeight
        {
            get => _defaultSillHeight;
            set => SetProperty(ref _defaultSillHeight, value);
        }

        private bool _isUpdated;
        public bool IsUpdated
        {
            get => _isUpdated;
            set => SetProperty(ref _isUpdated, value);
        }

        /// <summary>The Revit FamilySymbol ElementId for API operations.</summary>
        public ElementId? SymbolId { get; set; }
    }
}
