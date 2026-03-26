using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Pure WPF dialog that displays a filterable list of ViewSchedules.
    /// Supports single-select (default) and multi-select modes.
    /// </summary>
    public partial class SchedulePickerWindow : Window
    {
        private readonly List<ViewSchedule> _allSchedules;
        private readonly bool _multiSelect;

        /// <summary>
        /// The schedule the user selected in single-select mode (null if cancelled).
        /// </summary>
        public ViewSchedule? SelectedSchedule { get; private set; }

        /// <summary>
        /// The schedules the user selected in multi-select mode.
        /// </summary>
        public List<ViewSchedule> SelectedSchedules { get; private set; } = new();

        /// <summary>
        /// When true, export all selected schedules into one Excel file (multiple sheets).
        /// When false, export each schedule to a separate file.
        /// Only relevant in multi-select mode.
        /// </summary>
        public bool ExportAsSingleFile { get; private set; } = true;

        public SchedulePickerWindow(
            List<ViewSchedule> schedules,
            string? title = null,
            string? instruction = null,
            bool multiSelect = false)
        {
            InitializeComponent();
            _allSchedules = schedules.OrderBy(s => s.Name).ToList();
            _multiSelect = multiSelect;

            ScheduleList.ItemsSource = _allSchedules;

            if (_multiSelect)
            {
                ScheduleList.SelectionMode = SelectionMode.Extended;
                ExportModeGroup.Visibility = System.Windows.Visibility.Visible;
                InstructionText.Text = instruction ?? "Select schedules (Ctrl+Click for multiple):";
            }

            if (!string.IsNullOrEmpty(title))
                Title = title;
            if (!string.IsNullOrEmpty(instruction))
                InstructionText.Text = instruction;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string? filter = SearchBox.Text?.Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(filter))
            {
                ScheduleList.ItemsSource = _allSchedules;
            }
            else
            {
                ScheduleList.ItemsSource = _allSchedules
                    .Where(s => s.Name.ToLowerInvariant().Contains(filter))
                    .ToList();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_multiSelect)
            {
                SelectedSchedules = ScheduleList.SelectedItems
                    .Cast<ViewSchedule>()
                    .ToList();

                if (SelectedSchedules.Count == 0)
                    return;

                ExportAsSingleFile = SingleFileRadio.IsChecked == true;
                SelectedSchedule = SelectedSchedules[0];
            }
            else
            {
                SelectedSchedule = ScheduleList.SelectedItem as ViewSchedule;
                if (SelectedSchedule == null)
                    return;

                SelectedSchedules = new List<ViewSchedule> { SelectedSchedule };
            }

            DialogResult = true;
        }

        private void ScheduleList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_multiSelect)
                return; // Double-click only confirms in single-select mode

            SelectedSchedule = ScheduleList.SelectedItem as ViewSchedule;
            if (SelectedSchedule != null)
            {
                SelectedSchedules = new List<ViewSchedule> { SelectedSchedule };
                DialogResult = true;
            }
        }
    }
}
