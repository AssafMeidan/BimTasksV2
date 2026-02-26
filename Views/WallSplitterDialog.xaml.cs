using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Dialog for splitting a compound wall into individual walls per layer.
    /// Shows layer breakdown, type assignments, hosted element handling, and scope.
    /// </summary>
    public partial class WallSplitterDialog : Window
    {
        /// <summary>True if the user clicked Split (not Cancel).</summary>
        public bool Accepted { get; private set; }

        /// <summary>Original wall type name for display.</summary>
        public string WallTypeName { get; set; } = "";

        /// <summary>Original wall total width for display.</summary>
        public string WallWidth { get; set; } = "";

        /// <summary>Formatted info text shown in the top bar.</summary>
        public string InfoText { get; set; } = "";

        /// <summary>Layer rows displayed in the DataGrid.</summary>
        public ObservableCollection<LayerRowItem> LayerRows { get; } = new();

        /// <summary>Number of hosted elements found on the original wall.</summary>
        public int HostedElementCount { get; set; }

        /// <summary>
        /// Which layer index to host elements on. Null means auto (thickest/structural).
        /// </summary>
        public int? SelectedHostLayerIndex { get; set; }

        /// <summary>
        /// True to split all walls of this type in the project; false for selected only.
        /// </summary>
        public bool SplitAllOfType { get; set; }

        public WallSplitterDialog(string wallTypeName, double widthMm,
            List<LayerRowItem> rows, int hostedCount)
        {
            InitializeComponent();

            WallTypeName = wallTypeName;
            WallWidth = $"{widthMm:F0}";
            InfoText = $"Original: {wallTypeName}  Width: {widthMm:F0}mm";
            HostedElementCount = hostedCount;

            foreach (var row in rows)
                LayerRows.Add(row);

            // Build hosted element radio buttons programmatically
            if (hostedCount > 0)
            {
                HostedElementsGroup.Visibility = System.Windows.Visibility.Visible;
                HostedElementsHeader.Text = $"Hosted Elements ({hostedCount} found):";

                // Auto option (default)
                var autoRadio = new RadioButton
                {
                    Content = "Auto (thickest/structural)",
                    GroupName = "HostTarget",
                    IsChecked = true,
                    Margin = new Thickness(0, 2, 0, 2)
                };
                autoRadio.Checked += (_, _) => SelectedHostLayerIndex = null;
                HostedRadioPanel.Children.Add(autoRadio);

                // One radio per layer
                foreach (var row in rows)
                {
                    var layerRadio = new RadioButton
                    {
                        Content = $"Layer {row.Index}: {row.MaterialName} ({row.ThicknessMm}mm)",
                        GroupName = "HostTarget",
                        IsChecked = false,
                        Margin = new Thickness(0, 2, 0, 2),
                        Tag = row.Index
                    };
                    var capturedIndex = row.Index;
                    layerRadio.Checked += (_, _) => SelectedHostLayerIndex = capturedIndex;
                    HostedRadioPanel.Children.Add(layerRadio);
                }
            }

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

    /// <summary>
    /// Represents a single layer row in the wall splitter dialog.
    /// </summary>
    public class LayerRowItem
    {
        public int Index { get; set; }
        public string FunctionName { get; set; } = "";
        public string MaterialName { get; set; } = "";
        public string ThicknessMm { get; set; } = "";
        public List<WallTypeOption> AvailableTypes { get; set; } = new();
        public WallTypeOption? SelectedType { get; set; }
    }

    /// <summary>
    /// Represents a wall type option in the type assignment combo box.
    /// </summary>
    public class WallTypeOption
    {
        /// <summary>Display name: "TypeName" or "* TypeName" if auto-create.</summary>
        public string DisplayName { get; set; } = "";

        /// <summary>The existing WallType, or null if this is an auto-create option.</summary>
        public WallType? WallType { get; set; }

        /// <summary>True if this type will be auto-created during the split.</summary>
        public bool IsAutoCreate { get; set; }

        public override string ToString() => DisplayName;
    }
}
