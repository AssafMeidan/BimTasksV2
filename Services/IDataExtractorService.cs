using System.Collections.Generic;
using Autodesk.Revit.DB;
using BimTasksV2.Models;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Extracts model element data for BIM Israel web application export.
    /// Output JSON is compatible with the web app's upload route.
    /// </summary>
    public interface IDataExtractorService
    {
        /// <summary>
        /// Extracts all elements from the document for web app upload.
        /// Filters by the specified phase using "Show Complete" logic (New + Existing elements only).
        /// </summary>
        /// <param name="doc">The Revit document to extract from.</param>
        /// <param name="phaseId">The active view's phase ElementId for filtering.</param>
        /// <returns>List of BimElementDto for JSON serialization.</returns>
        List<BimElementDto> ExtractElements(Document doc, ElementId phaseId);

        /// <summary>
        /// Generates a validation summary for cross-checking export against Revit schedules.
        /// Stair runs and landings are excluded from schedule comparison (calculated values).
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="extractedData">The extracted element DTOs.</param>
        /// <returns>Validation summary with category/family breakdowns.</returns>
        ExportValidationSummary GenerateValidationSummary(Document doc, List<BimElementDto> extractedData);
    }
}
