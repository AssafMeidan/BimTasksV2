using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Models;
using BimTasksV2.Services;
using Microsoft.Win32;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Exports model elements to JSON for BIM Israel web application upload.
    /// Orchestrates: validate -> extract -> serialize -> compare -> show results.
    /// Uses System.Text.Json for serialization (not Newtonsoft.Json).
    /// Uses fully-qualified BimTasksV2.Infrastructure.ContainerLocator because the
    /// Commands namespace has its own Infrastructure sub-namespace.
    /// </summary>
    public class WebAppDataExtractorHandler : ICommandHandler
    {
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        public void Execute(UIApplication uiApp)
        {
            try
            {
                Document doc = uiApp.ActiveUIDocument.Document;

                // 1. Resolve services from DI
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var extractorService = container.Resolve<IDataExtractorService>();
                var scheduleService = container.Resolve<IValidationScheduleService>();

                // 2. Create validation schedules if they don't exist
                var createdSchedules = scheduleService.EnsureSchedules(doc);

                // 3. Get save location from user
                string filePath = GetSaveFilePath(doc.Title);
                if (string.IsNullOrEmpty(filePath))
                    return; // User cancelled

                // 4. Extract elements (filtered by active view's phase)
                var activeView = uiApp.ActiveUIDocument.ActiveView;
                var phaseId = activeView.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsElementId() ?? ElementId.InvalidElementId;
                var elementDtos = extractorService.ExtractElements(doc, phaseId);

                if (elementDtos.Count == 0)
                {
                    TaskDialog.Show("Export", "No elements found to export.\n\nMake sure the model contains physical elements with geometry.");
                    return;
                }

                // 5. Generate validation summary
                var validationSummary = extractorService.GenerateValidationSummary(doc, elementDtos);

                // 6. Save elements JSON
                string jsonOutput = JsonSerializer.Serialize(elementDtos, JsonOptions);
                File.WriteAllText(filePath, jsonOutput, new UTF8Encoding(false));

                // 7. Save validation summary as separate file
                string validationFilePath = Path.Combine(
                    Path.GetDirectoryName(filePath),
                    Path.GetFileNameWithoutExtension(filePath) + "_validation.json");
                string validationJson = JsonSerializer.Serialize(validationSummary, JsonOptions);
                File.WriteAllText(validationFilePath, validationJson, new UTF8Encoding(false));

                // 8. Build phase filter for validation comparison
                HashSet<long> validCreationPhaseIds = null;
                if (phaseId != null && phaseId != ElementId.InvalidElementId)
                {
                    validCreationPhaseIds = new HashSet<long>();
                    foreach (Phase p in doc.Phases)
                    {
                        validCreationPhaseIds.Add(p.Id.Value);
                        if (p.Id.Value == phaseId.Value)
                            break;
                    }
                }

                // 9. Compare export totals against Revit schedule totals
                var discrepancies = CompareAgainstSchedules(doc, scheduleService, validationSummary, validCreationPhaseIds);
                bool hasDiscrepancies = discrepancies.Count > 0;

                // 10. Show results
                var stats = GenerateExportStats(elementDtos);
                string scheduleInfo = createdSchedules.Count > 0
                    ? $"\n\nCreated {createdSchedules.Count} validation schedules (BIM_Validation_*)"
                    : "\n\nValidation schedules already exist";

                string validationInfo = GenerateValidationInfo(validationSummary);

                string resultHeader;
                string discrepancyInfo;
                TaskDialogIcon icon;

                if (hasDiscrepancies)
                {
                    int categoryCount = discrepancies.Count(d => d.Contains("[") && d.Contains("]"));
                    resultHeader = $"VALIDATION FAILED\n{categoryCount} categories with discrepancies - review needed";
                    discrepancyInfo = "\n\nDISCREPANCIES:" + string.Join("\n", discrepancies);
                    icon = TaskDialogIcon.TaskDialogIconError;
                }
                else
                {
                    resultHeader = $"VALIDATION PASSED\nAll {elementDtos.Count:N0} elements verified successfully";
                    discrepancyInfo = "\n\nAll export totals match Revit schedules";
                    icon = TaskDialogIcon.TaskDialogIconInformation;
                }

                TaskDialog td = new TaskDialog("BIM Israel Export")
                {
                    MainInstruction = resultHeader,
                    MainContent = discrepancyInfo + scheduleInfo + validationInfo +
                        $"\n\nFiles saved:" +
                        $"\n   Elements: {Path.GetFileName(filePath)}" +
                        $"\n   Validation: {Path.GetFileName(validationFilePath)}" +
                        $"\n\n" + stats,
                    MainIcon = icon
                };
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open containing folder");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Close");

                TaskDialogResult result = td.Show();
                if (result == TaskDialogResult.CommandLink1)
                {
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                }

                Log.Information("WebApp export completed: {Count} elements to {Path}", elementDtos.Count, filePath);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebApp export failed");
                TaskDialog.Show("Export Error", $"Failed to export:\n\n{ex.Message}\n\n{ex.StackTrace}");
            }
        }

        #region Save File Dialog

        private string GetSaveFilePath(string documentTitle)
        {
            string cleanTitle = string.IsNullOrEmpty(documentTitle)
                ? "BimExport"
                : CleanFileName(documentTitle);

            string defaultFileName = $"{cleanTitle}_{DateTime.Now:yyyyMMdd_HHmmss}.json";

            var dialog = new SaveFileDialog
            {
                Title = "Export BIM Data for Web Application",
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                DefaultExt = "json",
                FileName = defaultFileName,
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            bool? dialogResult = dialog.ShowDialog();
            return dialogResult == true ? dialog.FileName : null;
        }

        private string CleanFileName(string fileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string clean = fileName;
            foreach (char c in invalidChars)
            {
                clean = clean.Replace(c, '_');
            }
            if (clean.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
            {
                clean = clean.Substring(0, clean.Length - 4);
            }
            return clean;
        }

        #endregion Save File Dialog

        #region Statistics & Validation

        private string GenerateExportStats(List<BimElementDto> elements)
        {
            var stats = new StringBuilder();
            var categoryGroups = new Dictionary<string, int>();
            int withVolume = 0, withArea = 0, withLength = 0;
            int withBoqCode = 0, withZone = 0, withWorkStage = 0;
            int withSeifChoze = 0, withSubcontractor = 0;
            var levels = new HashSet<string>();

            foreach (var e in elements)
            {
                if (!categoryGroups.ContainsKey(e.Category))
                    categoryGroups[e.Category] = 0;
                categoryGroups[e.Category]++;

                if (e.Volume > 0) withVolume++;
                if (e.Area > 0) withArea++;
                if (e.Length > 0) withLength++;

                if (!string.IsNullOrEmpty(e.BoqCode)) withBoqCode++;
                if (!string.IsNullOrEmpty(e.Zone)) withZone++;
                if (!string.IsNullOrEmpty(e.WorkStage)) withWorkStage++;

                if (!string.IsNullOrEmpty(e.SeifChoze)) withSeifChoze++;
                if (!string.IsNullOrEmpty(e.SubcontractorName)) withSubcontractor++;

                if (!string.IsNullOrEmpty(e.Level) && e.Level != "Unassigned")
                    levels.Add(e.Level);
            }

            stats.AppendLine("Export Statistics:");
            stats.AppendLine($"   Categories: {categoryGroups.Count}");
            stats.AppendLine($"   Levels: {levels.Count}");
            stats.AppendLine();
            stats.AppendLine("Quantities:");
            stats.AppendLine($"   With Volume: {withVolume:N0}");
            stats.AppendLine($"   With Area: {withArea:N0}");
            stats.AppendLine($"   With Length: {withLength:N0}");
            stats.AppendLine();
            stats.AppendLine("BI_ Commercial Data:");
            stats.AppendLine($"   With BOQ Code: {withBoqCode:N0}");
            stats.AppendLine($"   With Zone: {withZone:N0}");
            stats.AppendLine($"   With Work Stage: {withWorkStage:N0}");
            stats.AppendLine();
            stats.AppendLine("Legacy BOQ Data:");
            stats.AppendLine($"   With SeifChoze: {withSeifChoze:N0}");
            stats.AppendLine($"   With Subcontractor: {withSubcontractor:N0}");

            return stats.ToString();
        }

        private string GenerateValidationInfo(ExportValidationSummary validation)
        {
            var info = new StringBuilder();
            info.AppendLine();
            info.AppendLine("Validation Summary:");
            info.AppendLine($"   Schedulable elements: {validation.TotalElements - validation.ElementsExcludedFromComparison:N0}");
            info.AppendLine($"   (Compare against BIM_Validation_* schedules)");

            if (validation.ElementsExcludedFromComparison > 0)
            {
                info.AppendLine();
                info.AppendLine("Excluded from comparison (calculated values):");
                info.AppendLine($"   Stair Runs: {validation.StairBreakdown.StairRunCount} (TreadLength: {validation.StairBreakdown.TotalTreadLength:F2} m)");
                info.AppendLine($"   Landings: {validation.StairBreakdown.LandingCount} (Area: {validation.StairBreakdown.TotalLandingArea:F2} m2)");
                info.AppendLine($"   -> Verify using 'runCalculations' in JSON");
            }

            return info.ToString();
        }

        #endregion Statistics & Validation

        #region Schedule Comparison

        /// <summary>
        /// Compares export validation totals against Revit schedule totals.
        /// </summary>
        private List<string> CompareAgainstSchedules(
            Document doc,
            IValidationScheduleService scheduleService,
            ExportValidationSummary validation,
            HashSet<long> validCreationPhaseIds = null)
        {
            var discrepancies = new List<string>();
            const double tolerance = 0.1;

            try
            {
                var scheduleTotals = scheduleService.GetTotals(doc, validCreationPhaseIds);

                var categoryMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Walls", "Walls" },
                    { "Floors", "Floors" },
                    { "Ceilings", "Ceilings" },
                    { "Roofs", "Roofs" },
                    { "Doors", "Doors" },
                    { "Windows", "Windows" },
                    { "Structural Columns", "StructuralColumns" },
                    { "Structural Framing", "StructuralFraming" },
                    { "Structural Foundations", "StructuralFoundation" },
                    { "Stairs", "Stairs" },
                    { "Railings", "Railings" },
                    { "Furniture", "Furniture" },
                    { "Generic Models", "GenericModels" },
                    { "Casework", "Casework" },
                    { "Mechanical Equipment", "MechanicalEquipment" },
                    { "Plumbing Fixtures", "PlumbingFixtures" },
                    { "Electrical Fixtures", "ElectricalFixtures" },
                    { "Electrical Equipment", "ElectricalEquipment" },
                    { "Lighting Fixtures", "LightingFixtures" },
                    { "Ducts", "Ducts" },
                    { "Pipes", "Pipes" },
                    { "Cable Trays", "CableTrays" },
                    { "Conduits", "Conduits" }
                };

                foreach (var kvp in validation.CountByCategory)
                {
                    string category = kvp.Key;
                    int exportCount = kvp.Value;

                    if (!categoryMappings.TryGetValue(category, out var schedName))
                        continue;

                    if (!scheduleTotals.TryGetValue(schedName, out var schedTotals))
                        continue;

                    // Skip if schedule appears empty
                    if (schedTotals.Count == 0 && schedTotals.Volume == 0 && schedTotals.Area == 0 && schedTotals.Length == 0)
                        continue;

                    double exportVolume = validation.VolumeByCategory.ContainsKey(category) ? validation.VolumeByCategory[category] : 0;
                    double exportArea = validation.AreaByCategory.ContainsKey(category) ? validation.AreaByCategory[category] : 0;
                    double exportLength = validation.LengthByCategory.ContainsKey(category) ? validation.LengthByCategory[category] : 0;

                    bool countMismatch = exportCount != schedTotals.Count;
                    bool volumeMismatch = schedTotals.Volume > 0 && Math.Abs(exportVolume - schedTotals.Volume) > tolerance;
                    bool areaMismatch = schedTotals.Area > 0 && Math.Abs(exportArea - schedTotals.Area) > tolerance;
                    bool lengthMismatch = schedTotals.Length > 0 && Math.Abs(exportLength - schedTotals.Length) > tolerance;

                    if (countMismatch || volumeMismatch || areaMismatch || lengthMismatch)
                    {
                        discrepancies.Add($"\n   [{category}]");
                        discrepancies.Add($"                    JSON      Revit");
                        discrepancies.Add($"   Count:      {exportCount,10}  {schedTotals.Count,10}{(countMismatch ? " <-" : "")}");

                        if (exportVolume > 0 || schedTotals.Volume > 0)
                            discrepancies.Add($"   Volume m3:  {exportVolume,10:F2}  {schedTotals.Volume,10:F2}{(volumeMismatch ? " <-" : "")}");

                        if (exportArea > 0 || schedTotals.Area > 0)
                            discrepancies.Add($"   Area m2:    {exportArea,10:F2}  {schedTotals.Area,10:F2}{(areaMismatch ? " <-" : "")}");

                        if (exportLength > 0 || schedTotals.Length > 0)
                            discrepancies.Add($"   Length m:   {exportLength,10:F2}  {schedTotals.Length,10:F2}{(lengthMismatch ? " <-" : "")}");

                        // Show family-level breakdown for failing families
                        var familyDiscrepancies = GetFamilyDiscrepancies(doc, scheduleService, validation, category, schedName, tolerance, validCreationPhaseIds);
                        if (familyDiscrepancies.Count > 0)
                        {
                            discrepancies.Add($"\n   Families with discrepancies:");
                            discrepancies.AddRange(familyDiscrepancies);
                        }
                    }
                }
            }
            catch
            {
                discrepancies.Add("   Could not compare against schedules (schedule read error)");
            }

            return discrepancies;
        }

        /// <summary>
        /// Gets family-level discrepancies for a category that failed validation.
        /// </summary>
        private List<string> GetFamilyDiscrepancies(
            Document doc,
            IValidationScheduleService scheduleService,
            ExportValidationSummary validation,
            string categoryName,
            string scheduleName,
            double tolerance,
            HashSet<long> validCreationPhaseIds = null)
        {
            const int MaxFamiliesToShow = 5;
            var familyDiscrepancies = new List<string>();
            int discrepancyCount = 0;
            int totalDiscrepancies = 0;

            try
            {
                var revitFamilyTotals = scheduleService.GetFamilyTotals(doc, scheduleName, validCreationPhaseIds);

                var jsonFamilyTotals = validation.TotalsByFamily
                    .Where(kvp => kvp.Value.Category == categoryName)
                    .ToDictionary(kvp => kvp.Value.FamilyName, kvp => kvp.Value);

                var allFamilies = new HashSet<string>(revitFamilyTotals.Keys);
                foreach (var jsonFamily in jsonFamilyTotals.Keys)
                    allFamilies.Add(jsonFamily);

                foreach (var familyName in allFamilies.OrderBy(f => f))
                {
                    var hasRevit = revitFamilyTotals.TryGetValue(familyName, out var revitTotals);
                    var hasJson = jsonFamilyTotals.TryGetValue(familyName, out var jsonTotals);

                    int jsonCount = hasJson ? jsonTotals.Count : 0;
                    int revitCount = hasRevit ? revitTotals.Count : 0;
                    double jsonVol = hasJson ? jsonTotals.Volume : 0;
                    double revitVol = hasRevit ? revitTotals.Volume : 0;
                    double jsonArea = hasJson ? jsonTotals.Area : 0;
                    double revitArea = hasRevit ? revitTotals.Area : 0;
                    double jsonLen = hasJson ? jsonTotals.Length : 0;
                    double revitLen = hasRevit ? revitTotals.Length : 0;

                    bool countMismatch = jsonCount != revitCount;
                    bool volumeMismatch = (revitVol > 0 || jsonVol > 0) && Math.Abs(jsonVol - revitVol) > tolerance;
                    bool areaMismatch = (revitArea > 0 || jsonArea > 0) && Math.Abs(jsonArea - revitArea) > tolerance;
                    bool lengthMismatch = (revitLen > 0 || jsonLen > 0) && Math.Abs(jsonLen - revitLen) > tolerance;

                    if (countMismatch || volumeMismatch || areaMismatch || lengthMismatch)
                    {
                        totalDiscrepancies++;

                        if (discrepancyCount >= MaxFamiliesToShow)
                            continue;

                        discrepancyCount++;
                        familyDiscrepancies.Add($"\n      - {familyName}");
                        familyDiscrepancies.Add($"                       JSON      Revit");

                        if (countMismatch || jsonCount > 0 || revitCount > 0)
                            familyDiscrepancies.Add($"        Count:     {jsonCount,10}  {revitCount,10}{(countMismatch ? " <-" : "")}");

                        if (volumeMismatch || jsonVol > 0 || revitVol > 0)
                            familyDiscrepancies.Add($"        Volume m3: {jsonVol,10:F2}  {revitVol,10:F2}{(volumeMismatch ? " <-" : "")}");

                        if (areaMismatch || jsonArea > 0 || revitArea > 0)
                            familyDiscrepancies.Add($"        Area m2:   {jsonArea,10:F2}  {revitArea,10:F2}{(areaMismatch ? " <-" : "")}");

                        if (lengthMismatch || jsonLen > 0 || revitLen > 0)
                            familyDiscrepancies.Add($"        Length m:  {jsonLen,10:F2}  {revitLen,10:F2}{(lengthMismatch ? " <-" : "")}");
                    }
                }

                if (totalDiscrepancies > MaxFamiliesToShow)
                {
                    familyDiscrepancies.Add($"\n      ... and {totalDiscrepancies - MaxFamiliesToShow} more families");
                    familyDiscrepancies.Add($"      (See *_validation.json for full details)");
                }
            }
            catch
            {
                // If family comparison fails, just skip it
            }

            return familyDiscrepancies;
        }

        #endregion Schedule Comparison
    }
}
