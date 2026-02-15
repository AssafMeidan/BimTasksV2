using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using BimTasksV2.Models;
using Prism.Commands;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Dialog for selecting a window family and editing/duplicating its types.
    /// The handler provides window families; this dialog provides selection UI.
    /// </summary>
    public partial class CreateWindowFamiliesDialog : Window
    {
        /// <summary>True if the user clicked OK.</summary>
        public bool Accepted { get; private set; }

        /// <summary>The currently selected window family model.</summary>
        public WindowFamilyModel? SelectedFamily
        {
            get => _selectedFamily;
            set
            {
                _selectedFamily = value;
                WindowSymbols.Clear();
                if (value != null)
                {
                    foreach (var sym in value.Symbols)
                        WindowSymbols.Add(sym);
                    SelectedFamilyName = value.Name;
                }
                else
                {
                    SelectedFamilyName = "";
                }
            }
        }
        private WindowFamilyModel? _selectedFamily;

        /// <summary>Name of the selected family for display.</summary>
        public string SelectedFamilyName { get; set; } = "";

        /// <summary>All available window families.</summary>
        public ObservableCollection<WindowFamilyModel> WindowFamilies { get; } = new();

        /// <summary>Symbols (types) of the selected family.</summary>
        public ObservableCollection<WindowSymbolModel> WindowSymbols { get; } = new();

        /// <summary>Command to duplicate the currently selected type.</summary>
        public DelegateCommand DuplicateTypeCommand { get; }

        public CreateWindowFamiliesDialog(List<WindowFamilyModel> families)
        {
            InitializeComponent();

            DuplicateTypeCommand = new DelegateCommand(OnDuplicateType, () => SelectedFamily != null);

            foreach (var fam in families)
                WindowFamilies.Add(fam);

            if (WindowFamilies.Count > 0)
                SelectedFamily = WindowFamilies[0];

            DataContext = this;
        }

        private void OnDuplicateType()
        {
            // Stub: the handler does the actual duplication via Revit API
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Accepted = true;
            DialogResult = true;
            Close();
        }
    }
}
