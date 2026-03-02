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
        public double OverlapPercent { get; set; }
        public string OverlapPercentText { get; set; } = "";
        public string TypeName { get; set; } = "";
        public int WallCount { get; set; }
        public string WallIds { get; set; } = "";
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

            // Sort by overlap percentage descending — exact duplicates (100%) first
            var sorted = groups.OrderByDescending(g => g.OverlapPercent).ToList();

            int totalWalls = sorted.Sum(g => g.Walls.Count);
            SummaryText = $"Found {sorted.Count} overlapping group(s) containing {totalWalls} walls";

            for (int i = 0; i < sorted.Count; i++)
            {
                var g = sorted[i];
                Groups.Add(new OverlappingGroupRow
                {
                    Index = i + 1,
                    OverlapPercent = g.OverlapPercent,
                    OverlapPercentText = $"{g.OverlapPercent:F0}%",
                    TypeName = g.WallType.Name,
                    WallCount = g.Walls.Count,
                    WallIds = string.Join(", ", g.Walls.Select(w => w.Id.Value)),
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
