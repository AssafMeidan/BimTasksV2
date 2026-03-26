using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using BimTasksV2.Commands.Handlers;

namespace BimTasksV2.Views
{
    public partial class FilterToLegendDialog : Window
    {
        public bool Accepted { get; private set; }
        public List<FilterLegendItem> SelectedFilters { get; private set; } = new();

        /// <summary>True = create new legend, False = update existing.</summary>
        public bool IsCreateNew { get; private set; } = true;

        /// <summary>The existing legend view info selected for update (null if creating new).</summary>
        public LegendViewInfo? SelectedExistingLegend { get; private set; }

        public ObservableCollection<FilterDialogRow> Rows { get; } = new();

        public FilterToLegendDialog(List<FilterLegendItem> filters, List<LegendViewInfo> existingLegends)
        {
            InitializeComponent();

            // Populate filter rows
            foreach (var f in filters)
            {
                Rows.Add(new FilterDialogRow
                {
                    FilterName = f.FilterName,
                    IsSelected = f.HasColorOverride,
                    DisplayColor = f.HasColorOverride
                        ? Color.FromRgb(f.OverrideColor.Red, f.OverrideColor.Green, f.OverrideColor.Blue)
                        : Color.FromRgb(200, 200, 200),
                    SourceItem = f
                });
            }

            FilterListBox.ItemsSource = Rows;

            // Populate existing legends dropdown
            ExistingLegendsCombo.ItemsSource = existingLegends;
            if (existingLegends.Count > 0)
                ExistingLegendsCombo.SelectedIndex = 0;

            // If no existing legends, disable the update option
            if (existingLegends.Count == 0)
                UpdateExistingRadio.IsEnabled = false;

            UpdateInfoText();
        }

        private void LegendMode_Changed(object sender, RoutedEventArgs e)
        {
            // Guard: this event fires during InitializeComponent before controls are assigned
            if (ActionButton == null || ExistingLegendsCombo == null) return;

            bool isUpdate = UpdateExistingRadio.IsChecked == true;
            ExistingLegendsCombo.IsEnabled = isUpdate;
            ActionButton.Content = isUpdate ? "Update Legend" : "Create Legend";
        }

        private void UpdateInfoText()
        {
            int selected = Rows.Count(r => r.IsSelected);
            InfoText.Text = $"{selected} of {Rows.Count} filters selected";
        }

        private void MoveUpButton_Click(object sender, RoutedEventArgs e)
        {
            int index = FilterListBox.SelectedIndex;
            if (index <= 0) return;

            var item = Rows[index];
            Rows.RemoveAt(index);
            Rows.Insert(index - 1, item);
            FilterListBox.SelectedIndex = index - 1;
        }

        private void MoveDownButton_Click(object sender, RoutedEventArgs e)
        {
            int index = FilterListBox.SelectedIndex;
            if (index < 0 || index >= Rows.Count - 1) return;

            var item = Rows[index];
            Rows.RemoveAt(index);
            Rows.Insert(index + 1, item);
            FilterListBox.SelectedIndex = index + 1;
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in Rows) row.IsSelected = true;
            FilterListBox.Items.Refresh();
            UpdateInfoText();
        }

        private void SelectNoneButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var row in Rows) row.IsSelected = false;
            FilterListBox.Items.Refresh();
            UpdateInfoText();
        }

        private void ActionButton_Click(object sender, RoutedEventArgs e)
        {
            // Collect selected filters in display order
            SelectedFilters = Rows
                .Where(r => r.IsSelected)
                .Select(r => r.SourceItem)
                .ToList();

            if (SelectedFilters.Count == 0)
            {
                MessageBox.Show("Please select at least one filter.", "BimTasksV2",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IsCreateNew = CreateNewRadio.IsChecked == true;

            if (!IsCreateNew)
            {
                SelectedExistingLegend = ExistingLegendsCombo.SelectedItem as LegendViewInfo;
                if (SelectedExistingLegend == null)
                {
                    MessageBox.Show("Please select a legend view to update.", "BimTasksV2",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

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

    public class FilterDialogRow
    {
        public string FilterName { get; set; } = "";
        public bool IsSelected { get; set; } = true;
        public Color DisplayColor { get; set; }
        public FilterLegendItem SourceItem { get; set; } = null!;
    }

    public class LegendViewInfo
    {
        public long ViewId { get; set; }
        public string Name { get; set; } = "";

        public override string ToString() => Name;
    }
}
