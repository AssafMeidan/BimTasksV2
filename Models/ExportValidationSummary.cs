using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace BimTasksV2.Models
{
    /// <summary>
    /// Validation summary for cross-checking export against Revit schedules.
    /// Stair runs (tread length) and landings (area) are excluded from schedule comparison
    /// because these are calculated values not directly available in Revit schedules.
    /// Uses System.Text.Json attributes (not Newtonsoft.Json).
    /// </summary>
    public class ExportValidationSummary
    {
        #region Category Totals (Compare against BIM_Validation_* schedule grand totals)

        /// <summary>
        /// Total element count by category (excludes Stair Runs and Landings).
        /// Compare against: BIM_Validation_* schedule Grand Total count.
        /// </summary>
        [JsonPropertyName("countByCategory")]
        public Dictionary<string, int> CountByCategory { get; set; } = new Dictionary<string, int>();

        /// <summary>
        /// Total volume by category in m3.
        /// Compare against: BIM_Validation_* schedule Grand Total volume.
        /// </summary>
        [JsonPropertyName("volumeByCategory_m3")]
        public Dictionary<string, double> VolumeByCategory { get; set; } = new Dictionary<string, double>();

        /// <summary>
        /// Total area by category in m2.
        /// Compare against: BIM_Validation_* schedule Grand Total area.
        /// </summary>
        [JsonPropertyName("areaByCategory_m2")]
        public Dictionary<string, double> AreaByCategory { get; set; } = new Dictionary<string, double>();

        /// <summary>
        /// Total length by category in m.
        /// Compare against: BIM_Validation_* schedule Grand Total length.
        /// </summary>
        [JsonPropertyName("lengthByCategory_m")]
        public Dictionary<string, double> LengthByCategory { get; set; } = new Dictionary<string, double>();

        #endregion

        #region Family Totals (Compare against BIM_Validation_* schedule family subtotals)

        /// <summary>
        /// Totals broken down by Category > Family.
        /// Key format: "Category|FamilyName"
        /// Compare against: BIM_Validation_* schedule subtotals per Family and Type.
        /// </summary>
        [JsonPropertyName("totalsByFamily")]
        public Dictionary<string, FamilyTotals> TotalsByFamily { get; set; } = new Dictionary<string, FamilyTotals>();

        #endregion

        #region Stair Breakdown (Manual verification only - not schedulable)

        /// <summary>
        /// Stair calculation details for manual spot-checking.
        /// These values cannot be verified against Revit schedules.
        /// </summary>
        [JsonPropertyName("stairBreakdown")]
        public StairValidationBreakdown StairBreakdown { get; set; } = new StairValidationBreakdown();

        #endregion

        #region Metadata

        /// <summary>Total elements exported</summary>
        [JsonPropertyName("totalElements")]
        public int TotalElements { get; set; }

        /// <summary>Elements excluded from schedule comparison (stair runs + landings)</summary>
        [JsonPropertyName("elementsExcludedFromComparison")]
        public int ElementsExcludedFromComparison { get; set; }

        /// <summary>Warnings for manual review</summary>
        [JsonPropertyName("warnings")]
        public List<string> Warnings { get; set; } = new List<string>();

        /// <summary>Export timestamp</summary>
        [JsonPropertyName("exportTimestamp")]
        public string ExportTimestamp { get; set; }

        #endregion
    }

    /// <summary>
    /// Stair-specific breakdown for manual verification.
    /// These are calculated values that cannot be compared against Revit schedules.
    /// </summary>
    public class StairValidationBreakdown
    {
        /// <summary>Number of parent stair elements</summary>
        [JsonPropertyName("parentStairCount")]
        public int ParentStairCount { get; set; }

        /// <summary>Total concrete volume from parent stairs (m3)</summary>
        [JsonPropertyName("parentStairVolume_m3")]
        public double ParentStairVolume { get; set; }

        /// <summary>Number of stair run elements</summary>
        [JsonPropertyName("stairRunCount")]
        public int StairRunCount { get; set; }

        /// <summary>
        /// Total tread length (m) - CALCULATED as TreadDepth x NumRisers per run.
        /// Manual verification: Check a few runs against (Actual Tread Depth x Actual Number of Risers).
        /// </summary>
        [JsonPropertyName("totalTreadLength_m")]
        public double TotalTreadLength { get; set; }

        /// <summary>Number of landing elements</summary>
        [JsonPropertyName("landingCount")]
        public int LandingCount { get; set; }

        /// <summary>
        /// Total landing area (m2) - from geometry or bounding box.
        /// Manual verification: Spot-check a few landings visually.
        /// </summary>
        [JsonPropertyName("totalLandingArea_m2")]
        public double TotalLandingArea { get; set; }

        /// <summary>
        /// Individual run calculations for spot-checking.
        /// Shows TreadDepth, RiserCount, and resulting TreadLength for each run.
        /// </summary>
        [JsonPropertyName("runCalculations")]
        public List<StairRunCalculation> RunCalculations { get; set; } = new List<StairRunCalculation>();
    }

    /// <summary>
    /// Individual stair run calculation details for verification.
    /// </summary>
    public class StairRunCalculation
    {
        [JsonPropertyName("elementId")]
        public string ElementId { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        /// <summary>Tread depth in meters (from parent stair or run)</summary>
        [JsonPropertyName("treadDepth_m")]
        public double TreadDepth { get; set; }

        /// <summary>Number of risers in this run</summary>
        [JsonPropertyName("riserCount")]
        public int RiserCount { get; set; }

        /// <summary>Calculated tread length: TreadDepth x RiserCount</summary>
        [JsonPropertyName("calculatedTreadLength_m")]
        public double CalculatedTreadLength { get; set; }

        /// <summary>How the tread depth was obtained</summary>
        [JsonPropertyName("treadDepthSource")]
        public string TreadDepthSource { get; set; }
    }

    /// <summary>
    /// Totals for a specific Category + Family combination.
    /// Used for comparing against Revit schedule subtotals.
    /// </summary>
    public class FamilyTotals
    {
        [JsonPropertyName("category")]
        public string Category { get; set; }

        [JsonPropertyName("familyName")]
        public string FamilyName { get; set; }

        [JsonPropertyName("count")]
        public int Count { get; set; }

        [JsonPropertyName("volume_m3")]
        public double Volume { get; set; }

        [JsonPropertyName("area_m2")]
        public double Area { get; set; }

        [JsonPropertyName("length_m")]
        public double Length { get; set; }
    }
}
