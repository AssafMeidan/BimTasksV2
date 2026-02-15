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
    /// The caller passes in a list of schedules; the user picks one and clicks OK.
    /// </summary>
    public partial class SchedulePickerWindow : Window
    {
        private readonly List<ViewSchedule> _allSchedules;

        /// <summary>
        /// The schedule the user selected (null if cancelled).
        /// </summary>
        public ViewSchedule? SelectedSchedule { get; private set; }

        public SchedulePickerWindow(
            List<ViewSchedule> schedules,
            string? title = null,
            string? instruction = null)
        {
            InitializeComponent();
            _allSchedules = schedules.OrderBy(s => s.Name).ToList();
            ScheduleList.ItemsSource = _allSchedules;

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
            SelectedSchedule = ScheduleList.SelectedItem as ViewSchedule;
            if (SelectedSchedule != null)
            {
                DialogResult = true;
            }
        }

        private void ScheduleList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            SelectedSchedule = ScheduleList.SelectedItem as ViewSchedule;
            if (SelectedSchedule != null)
            {
                DialogResult = true;
            }
        }
    }
}
