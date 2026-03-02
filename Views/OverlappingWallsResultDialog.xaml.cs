using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using BimTasksV2.Helpers.OverlappingWallDetector;

namespace BimTasksV2.Views
{
    public enum OverlappingWallAction
    {
        Highlight,
        Unjoin,
        DeleteDuplicates
    }

    public class OverlappingGroupRow
    {
        public int Index { get; set; }
        public string TypeName { get; set; } = "";
        public int WallCount { get; set; }
        public string WallIds { get; set; } = "";
        public string OverlapMeters { get; set; } = "";
    }

    public partial class OverlappingWallsResultDialog : Window
    {
        public bool Accepted { get; private set; }
        public OverlappingWallAction ChosenAction { get; private set; }

        public string SummaryText { get; set; } = "";
        public ObservableCollection<OverlappingGroupRow> Groups { get; } = new();

        public OverlappingWallsResultDialog(List<OverlappingWallGroup> groups, Document doc)
        {
            InitializeComponent();

            int totalWalls = groups.Sum(g => g.Walls.Count);
            SummaryText = $"Found {groups.Count} overlapping group(s) containing {totalWalls} walls";

            for (int i = 0; i < groups.Count; i++)
            {
                var g = groups[i];
                Groups.Add(new OverlappingGroupRow
                {
                    Index = i + 1,
                    TypeName = g.WallType.Name,
                    WallCount = g.Walls.Count,
                    WallIds = string.Join(", ", g.Walls.Select(w => w.Id.Value)),
                    OverlapMeters = $"{g.OverlapLength * 0.3048:F2}"
                });
            }

            DataContext = this;
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            if (RbHighlight.IsChecked == true)
                ChosenAction = OverlappingWallAction.Highlight;
            else if (RbUnjoin.IsChecked == true)
                ChosenAction = OverlappingWallAction.Unjoin;
            else if (RbDelete.IsChecked == true)
                ChosenAction = OverlappingWallAction.DeleteDuplicates;

            Accepted = true;
            DialogResult = true;
            Close();
        }
    }
}
