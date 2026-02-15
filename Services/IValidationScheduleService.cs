using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Creates and manages BIM_Validation_* schedules in Revit for cross-checking exported data.
    /// </summary>
    public interface IValidationScheduleService
    {
        /// <summary>
        /// Ensures all validation schedules exist in the document. Creates missing ones.
        /// Must be called within or will create its own transaction.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <returns>List of schedule names that were created.</returns>
        List<string> EnsureSchedules(Document doc);

        /// <summary>
        /// Gets totals from all BIM_Validation_* schedules for comparison with export data.
        /// Phase filter excludes demolished elements to match "Show Complete" behavior.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="validCreationPhaseIds">Set of valid creation phase IDs (ElementId.Value), or null for no filtering.</param>
        /// <returns>Dictionary keyed by category short name (e.g., "Walls") with totals.</returns>
        Dictionary<string, ValidationScheduleTotals> GetTotals(Document doc, HashSet<long> validCreationPhaseIds = null);

        /// <summary>
        /// Gets family-level totals from a specific validation schedule.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="scheduleName">Category short name (e.g., "Walls").</param>
        /// <param name="validCreationPhaseIds">Set of valid creation phase IDs, or null for no filtering.</param>
        /// <returns>Dictionary keyed by family name with totals.</returns>
        Dictionary<string, ValidationScheduleTotals> GetFamilyTotals(Document doc, string scheduleName, HashSet<long> validCreationPhaseIds = null);
    }

    /// <summary>
    /// Totals extracted from a validation schedule.
    /// </summary>
    public class ValidationScheduleTotals
    {
        public int Count { get; set; }
        public double Volume { get; set; }  // m3
        public double Area { get; set; }    // m2
        public double Length { get; set; }  // m
    }
}
