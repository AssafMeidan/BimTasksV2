using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Simple WPF dialog for selecting a FloorType and Level.
    /// Used by the DWG floor import commands.
    /// </summary>
    public partial class FloorTypeLevelSelector : Window
    {
        public FloorType? SelectedFloorType { get; private set; }
        public Level? SelectedLevel { get; private set; }

        public FloorTypeLevelSelector(Document doc, Level? defaultLevel = null)
        {
            InitializeComponent();

            var floorTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .OrderBy(ft => ft.Name)
                .ToList();

            FloorTypeList.ItemsSource = floorTypes;
            if (floorTypes.Count > 0)
                FloorTypeList.SelectedIndex = 0;

            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();

            LevelList.ItemsSource = levels;

            if (defaultLevel != null)
            {
                var match = levels.FirstOrDefault(l => l.Id == defaultLevel.Id);
                if (match != null)
                    LevelList.SelectedItem = match;
            }
            else if (levels.Count > 0)
            {
                LevelList.SelectedIndex = 0;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedFloorType = FloorTypeList.SelectedItem as FloorType;
            SelectedLevel = LevelList.SelectedItem as Level;

            if (SelectedFloorType == null)
            {
                MessageBox.Show("Please select a floor type.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (SelectedLevel == null)
            {
                MessageBox.Show("Please select a level.", "Selection Required",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
