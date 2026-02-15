using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Result object containing all calculated values for selected Revit elements.
    /// </summary>
    public class ElementCalculationResult
    {
        // Walls
        public int WallCount { get; set; }
        public double WallAreaM2 { get; set; }
        public double WallVolumeM3 { get; set; }

        // Floors
        public int FloorCount { get; set; }
        public double FloorAreaM2 { get; set; }
        public double FloorVolumeM3 { get; set; }

        // Beams
        public int BeamCount { get; set; }
        public double BeamVolumeM3 { get; set; }
        public double BeamLengthM { get; set; }

        // Columns
        public int ColumnCount { get; set; }
        public double ColumnVolumeM3 { get; set; }
        public double ColumnHeightM { get; set; }

        // Structural Foundations
        public int FoundationCount { get; set; }
        public double FoundationVolumeM3 { get; set; }
        public double FoundationAreaM2 { get; set; }

        // Stairs (parent - concrete volume)
        public int StairCount { get; set; }
        public int StairRiserCount { get; set; }
        public double StairVolumeM3 { get; set; }
        public double StairUndersideAreaM2 { get; set; }
        public double StairTreadLengthM { get; set; }

        // Stair Runs (tread flooring length)
        public int StairRunCount { get; set; }
        public double StairRunTreadLengthM { get; set; }

        // Landings (landing flooring area)
        public int LandingCount { get; set; }
        public double LandingAreaM2 { get; set; }
        public double LandingVolumeM3 { get; set; }

        // Railings
        public int RailingCount { get; set; }
        public double RailingLengthM { get; set; }

        // Doors
        public int DoorCount { get; set; }

        // Windows
        public int WindowCount { get; set; }

        // Material breakdown
        public Dictionary<string, double> VolumeByMaterial { get; } = new Dictionary<string, double>();

        // Totals
        public double TotalConcreteVolumeM3 =>
            WallVolumeM3 + FloorVolumeM3 + BeamVolumeM3 + ColumnVolumeM3 +
            FoundationVolumeM3 + StairVolumeM3 + LandingVolumeM3;

        public int TotalStructuralCount =>
            WallCount + FloorCount + BeamCount + ColumnCount + FoundationCount;

        public int TotalElementCount =>
            WallCount + FloorCount + BeamCount + ColumnCount + FoundationCount +
            (StairRiserCount > 0 ? 1 : 0) + LandingCount + RailingCount + DoorCount + WindowCount;

        // Metadata
        public int SelectedElementCount { get; set; }
        public long CalculationTimeMs { get; set; }
        public bool HasErrors { get; set; }
        public List<string> Errors { get; } = new List<string>();
        public Dictionary<string, int> CategoryBreakdown { get; } = new Dictionary<string, int>();
    }

    /// <summary>
    /// Service for calculating area and volume of selected Revit elements.
    /// </summary>
    public interface IElementCalculationService
    {
        ElementCalculationResult CalculateElements(Document doc, ICollection<ElementId> selectedIds);
    }
}
