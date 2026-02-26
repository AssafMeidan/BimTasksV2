using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using Autodesk.Revit.DB;

namespace BimTasksV2.Views
{
    public partial class FloorSplitterDialog : Window
    {
        public bool Accepted { get; private set; }
        public string InfoText { get; set; } = "";
        public ObservableCollection<FloorLayerRowItem> LayerRows { get; } = new();
        public bool SplitAllOfType { get; set; }

        public FloorSplitterDialog(string floorTypeName, double thicknessMm, List<FloorLayerRowItem> rows)
        {
            InitializeComponent();

            InfoText = $"Original: {floorTypeName}  Thickness: {thicknessMm:F0}mm";

            foreach (var row in rows)
                LayerRows.Add(row);

            DataContext = this;
        }

        private void SplitButton_Click(object sender, RoutedEventArgs e)
        {
            SplitAllOfType = ScopeAll.IsChecked == true;
            Accepted = true;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Accepted = false;
            DialogResult = false;
            Close();
        }
    }

    public class FloorLayerRowItem
    {
        public int Index { get; set; }
        public string FunctionName { get; set; } = "";
        public string MaterialName { get; set; } = "";
        public string ThicknessMm { get; set; } = "";
        public List<FloorTypeOption> AvailableTypes { get; set; } = new();
        public FloorTypeOption? SelectedType { get; set; }
    }

    public class FloorTypeOption
    {
        public string DisplayName { get; set; } = "";
        public FloorType? FloorType { get; set; }
        public bool IsAutoCreate { get; set; }
        public override string ToString() => DisplayName;
    }
}
