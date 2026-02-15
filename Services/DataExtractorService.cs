using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using BimTasksV2.Models;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Extracts model element data for BIM Israel web application.
    /// Output JSON is compatible with the web app's upload route.
    /// Migrated from OldApp DataExtractorForWebApp using Revit 2025 APIs.
    /// </summary>
    public class DataExtractorService : IDataExtractorService
    {
        #region Category Configuration

        /// <summary>
        /// Categories to extract - physical elements only.
        /// </summary>
        private static readonly List<BuiltInCategory> CategoriesToExtract = new List<BuiltInCategory>
        {
            // Architecture / Structure
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_Stairs,
            BuiltInCategory.OST_StairsRuns,
            BuiltInCategory.OST_StairsLandings,
            BuiltInCategory.OST_StairsRailing,
            BuiltInCategory.OST_Ramps,
            BuiltInCategory.OST_Railings,
            BuiltInCategory.OST_RailingTopRail,
            BuiltInCategory.OST_RailingHandRail,
            BuiltInCategory.OST_Reveals,
            BuiltInCategory.OST_CurtainWallPanels,
            BuiltInCategory.OST_CurtainWallMullions,

            // MEP - Mechanical
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctAccessory,

            // MEP - Piping
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_Sprinklers,

            // Electrical
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_LightingDevices,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_CableTrayFitting,
            BuiltInCategory.OST_ConduitFitting,
            BuiltInCategory.OST_FireAlarmDevices,

            // General
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_FurnitureSystems,
            BuiltInCategory.OST_Casework,
            BuiltInCategory.OST_SpecialityEquipment,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Parts
        };

        /// <summary>
        /// Stair categories bypass the zero-geometry filter because their
        /// measurements are calculated in post-processing, not from generic parameters.
        /// Uses ElementId.Value for Revit 2025.
        /// </summary>
        private static readonly HashSet<long> StairCategoryIds = new HashSet<long>
        {
            (long)BuiltInCategory.OST_Stairs,
            (long)BuiltInCategory.OST_StairsRuns,
            (long)BuiltInCategory.OST_StairsLandings
        };

        #endregion Category Configuration

        #region Main Extraction Method

        /// <inheritdoc/>
        public List<BimElementDto> ExtractElements(Document doc, ElementId phaseId)
        {
            var extractedData = new List<BimElementDto>();

            // Pre-cache levels for O(1) lookup
            var levelCache = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .GroupBy(l => l.Name)
                .ToDictionary(g => g.Key, g => g.First());

            var filter = new ElementMulticategoryFilter(CategoriesToExtract);

            var collector = new FilteredElementCollector(doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent();

            // Build phase order for "Show Complete" filtering
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

            // Build stair element dictionaries during first pass
            var stairElements = new Dictionary<string, Element>();
            var runElements = new Dictionary<string, Element>();
            var landingElements = new Dictionary<string, Element>();

            // Pre-calculate tread lengths for stair runs from parent stairs
            var runTreadLengths = CalculateAllRunTreadLengths(doc);

            foreach (Element e in collector)
            {
                if (e.Category == null) continue;

                // Phase filter: match "Show Complete" -- skip demolished or future elements
                if (validCreationPhaseIds != null)
                {
                    var createdParam = e.get_Parameter(BuiltInParameter.PHASE_CREATED);
                    var demolishedParam = e.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);

                    if (createdParam != null && createdParam.HasValue)
                    {
                        long createdId = createdParam.AsElementId().Value;
                        if (!validCreationPhaseIds.Contains(createdId))
                            continue;
                    }

                    if (demolishedParam != null && demolishedParam.HasValue)
                    {
                        long demolishedId = demolishedParam.AsElementId().Value;
                        if (demolishedId != -1 && validCreationPhaseIds.Contains(demolishedId))
                            continue;
                    }
                }

                long catId = e.Category.Id.Value;

                // Build parameter cache once per element for O(1) lookups
                var paramCache = BuildParamCache(e);

                // Get quantities
                double volume = GetVolumeRobust(e, paramCache);
                double area = GetAreaRobust(e, paramCache);
                double length = GetLengthRobust(e, paramCache);

                // Filter: Only export if has geometry
                // Stair elements bypass this -- their values are calculated in post-processing
                bool isStairElement = StairCategoryIds.Contains(catId);
                if (!isStairElement && volume <= 0 && area <= 0 && length <= 0) continue;

                // Track stair elements for post-processing
                if (catId == (long)BuiltInCategory.OST_Stairs)
                    stairElements[e.UniqueId] = e;
                else if (catId == (long)BuiltInCategory.OST_StairsRuns)
                    runElements[e.UniqueId] = e;
                else if (catId == (long)BuiltInCategory.OST_StairsLandings)
                    landingElements[e.UniqueId] = e;

                // Get element type for type-level parameters
                ElementType elementType = GetElementType(doc, e);
                var typeParamCache = elementType != null ? BuildParamCache(elementType) : null;

                // Get level info
                string levelName = GetLevelRobust(doc, e, paramCache);
                double? levelElevation = GetLevelElevation(doc, e, levelName, levelCache);

                var dto = new BimElementDto
                {
                    // Identification
                    ExternalId = e.UniqueId,
                    Name = GetElementName(e, paramCache),
                    HostElementId = GetHostElementId(e),

                    // Classification
                    Category = e.Category.Name,
                    FamilyName = elementType?.FamilyName ?? "",
                    TypeName = e.Name,
                    Function = GetFunction(e, paramCache),

                    // Location
                    Level = levelName,
                    LevelElevation = levelElevation,
                    Phase = GetPhaseRobust(doc, e, paramCache),

                    // Geometry
                    Volume = volume,
                    Area = area,
                    Length = length,
                    Unit = GetBIParameter(paramCache, typeParamCache, "BI_QtyBasis", "QtyBasis", "qtyBasis"),

                    // Type Information
                    AssemblyCode = GetAssemblyCode(typeParamCache, elementType),
                    AssemblyDescription = GetAssemblyDescription(doc, e, elementType),
                    TypeComments = GetTypeParameter(typeParamCache, elementType, "Type Comments", BuiltInParameter.ALL_MODEL_TYPE_COMMENTS),
                    TypeDescription = GetTypeParameter(typeParamCache, elementType, "Description", BuiltInParameter.ALL_MODEL_DESCRIPTION),
                    Keynote = GetTypeParameter(typeParamCache, elementType, "Keynote", BuiltInParameter.KEYNOTE_PARAM),
                    TypeModel = GetTypeModelParameter(elementType),

                    // BI_ Commercial Parameters
                    BoqCode = GetBIParameter(paramCache, typeParamCache, "BI_BOQ_Code", "BOQ_Code", "boqCode"),
                    Zone = GetBIParameter(paramCache, typeParamCache, "BI_Zone", "zone", "Zone"),
                    WorkStage = GetBIParameter(paramCache, typeParamCache, "BI_WorkStage", "WorkStage", "workStage"),
                    IsPayItem = GetBIBoolParameter(paramCache, typeParamCache, "BI_IsPayItem", "IsPayItem", "isPayItem"),
                    QtyBasis = GetBIParameter(paramCache, typeParamCache, "BI_QtyBasis", "QtyBasis", "qtyBasis"),
                    QtyOverride = GetBIDoubleParameter(paramCache, typeParamCache, "BI_QtyOverride", "QtyOverride", "qtyOverride"),
                    QtyMultiplier = GetBIDoubleParameter(paramCache, typeParamCache, "BI_QtyMultiplier", "QtyMultiplier", "qtyMultiplier"),
                    UnitPrice = GetBIDoubleParameter(paramCache, typeParamCache, "BI_UnitPrice", "UnitPrice"),
                    ExecutionPercentage = GetBIDoubleParameter(paramCache, typeParamCache, "BI_ExecPct_ToDate", "executionPercentage", "ExecutionPercentage"),
                    PaidPercentage = GetBIDoubleParameter(paramCache, typeParamCache, "BI_PaidPct_ToDate", "paidPercentage", "PaidPercentage"),
                    Note = GetBIParameter(paramCache, typeParamCache, "BI_Note", "Note", "note"),
                    SourceElementId = GetBIParameter(paramCache, typeParamCache, "BI_SourceElementId", "SourceElementId", "sourceElementId"),
                    ComponentRole = GetBIParameter(paramCache, typeParamCache, "BI_ComponentRole", "ComponentRole", "componentRole"),

                    // Legacy Fields
                    SeifChoze = GetBOQParameter(paramCache, typeParamCache, "SeifChoze", "Seif Choze"),
                    SeifChozeDescription = GetBOQParameter(paramCache, typeParamCache, "SeifChozeDescription", "Seif Choze Description"),
                    IsMedida = GetIsMedida(paramCache, typeParamCache),
                    SubcontractorName = GetBOQParameter(paramCache, typeParamCache, "SubcontractorName", "Subcontractor", "Name1"),
                    TotalContractPrice = GetTotalContractPrice(paramCache, typeParamCache)
                };

                extractedData.Add(dto);
            }

            // Post-process: enforce one element = one measurement for stairs
            foreach (var dto in extractedData)
            {
                if (stairElements.TryGetValue(dto.ExternalId, out var stairEl))
                {
                    // Parent stair: volume only (concrete)
                    dto.Area = 0;
                    dto.Length = 0;
                    if (dto.Volume <= 0)
                        dto.Volume = GetVolumeFromGeometry(stairEl);
                }

                if (runElements.TryGetValue(dto.ExternalId, out var runEl))
                {
                    // Runs: tread length only (horizontal, not slope)
                    dto.Volume = 0;
                    dto.Area = 0;

                    if (runTreadLengths.TryGetValue(runEl.Id.Value, out double calculatedLength) && calculatedLength > 0)
                    {
                        dto.Length = calculatedLength;
                    }
                    else
                    {
                        calculatedLength = CalculateRunTreadLength(runEl);
                        if (calculatedLength > 0)
                            dto.Length = calculatedLength;
                    }
                }

                if (landingElements.TryGetValue(dto.ExternalId, out var landingEl))
                {
                    // Landings: area only (landing flooring)
                    dto.Volume = 0;
                    dto.Length = 0;
                    if (dto.Area <= 0)
                    {
                        double calculatedArea = GetAreaFromGeometry(landingEl);
                        if (calculatedArea > 0)
                            dto.Area = calculatedArea;
                    }
                }
            }

            return extractedData;
        }

        /// <inheritdoc/>
        public ExportValidationSummary GenerateValidationSummary(Document doc, List<BimElementDto> extractedData)
        {
            var summary = new ExportValidationSummary
            {
                TotalElements = extractedData.Count,
                ExportTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };

            var excludedCategories = new HashSet<string> { "Stair Runs", "Stair Landings" };

            foreach (var dto in extractedData)
            {
                string category = dto.Category ?? "Unknown";
                bool isExcluded = excludedCategories.Contains(category);

                if (isExcluded)
                {
                    summary.ElementsExcludedFromComparison++;

                    if (category == "Stair Runs")
                    {
                        summary.StairBreakdown.StairRunCount++;
                        summary.StairBreakdown.TotalTreadLength += dto.Length;
                    }
                    else if (category == "Stair Landings")
                    {
                        summary.StairBreakdown.LandingCount++;
                        summary.StairBreakdown.TotalLandingArea += dto.Area;
                    }
                }
                else
                {
                    if (!summary.CountByCategory.ContainsKey(category))
                    {
                        summary.CountByCategory[category] = 0;
                        summary.VolumeByCategory[category] = 0;
                        summary.AreaByCategory[category] = 0;
                        summary.LengthByCategory[category] = 0;
                    }

                    summary.CountByCategory[category]++;

                    if (dto.Volume > 0)
                        summary.VolumeByCategory[category] += dto.Volume;
                    if (dto.Area > 0)
                        summary.AreaByCategory[category] += dto.Area;
                    if (dto.Length > 0)
                        summary.LengthByCategory[category] += dto.Length;

                    // Track by Family
                    string familyName = dto.FamilyName ?? "Unknown";
                    string familyKey = $"{category}|{familyName}";
                    if (!summary.TotalsByFamily.ContainsKey(familyKey))
                    {
                        summary.TotalsByFamily[familyKey] = new FamilyTotals
                        {
                            Category = category,
                            FamilyName = familyName
                        };
                    }
                    var familyTotals = summary.TotalsByFamily[familyKey];
                    familyTotals.Count++;
                    if (dto.Volume > 0) familyTotals.Volume += dto.Volume;
                    if (dto.Area > 0) familyTotals.Area += dto.Area;
                    if (dto.Length > 0) familyTotals.Length += dto.Length;

                    // Track parent stairs
                    if (category == "Stairs")
                    {
                        summary.StairBreakdown.ParentStairCount++;
                        summary.StairBreakdown.ParentStairVolume += dto.Volume;
                    }
                }
            }

            // Round all category values
            foreach (var key in new List<string>(summary.VolumeByCategory.Keys))
                summary.VolumeByCategory[key] = Math.Round(summary.VolumeByCategory[key], 2);
            foreach (var key in new List<string>(summary.AreaByCategory.Keys))
                summary.AreaByCategory[key] = Math.Round(summary.AreaByCategory[key], 2);
            foreach (var key in new List<string>(summary.LengthByCategory.Keys))
                summary.LengthByCategory[key] = Math.Round(summary.LengthByCategory[key], 2);

            // Round all family values
            foreach (var familyTotals in summary.TotalsByFamily.Values)
            {
                familyTotals.Volume = Math.Round(familyTotals.Volume, 2);
                familyTotals.Area = Math.Round(familyTotals.Area, 2);
                familyTotals.Length = Math.Round(familyTotals.Length, 2);
            }

            summary.StairBreakdown.ParentStairVolume = Math.Round(summary.StairBreakdown.ParentStairVolume, 2);
            summary.StairBreakdown.TotalTreadLength = Math.Round(summary.StairBreakdown.TotalTreadLength, 2);
            summary.StairBreakdown.TotalLandingArea = Math.Round(summary.StairBreakdown.TotalLandingArea, 2);

            // Generate stair run calculations for spot-checking
            summary.StairBreakdown.RunCalculations = GenerateStairRunCalculations(doc);

            // Add verification instructions as warnings
            if (summary.StairBreakdown.StairRunCount > 0 || summary.StairBreakdown.LandingCount > 0)
            {
                summary.Warnings.Add(
                    $"EXCLUDED FROM COMPARISON: {summary.StairBreakdown.StairRunCount} stair runs (tread length) " +
                    $"and {summary.StairBreakdown.LandingCount} landings (area) - these are calculated values, " +
                    "verify manually using runCalculations breakdown");
            }

            return summary;
        }

        #endregion Main Extraction Method

        #region Stair Helpers

        /// <summary>
        /// Pre-calculates tread lengths for all stair runs by accessing parent Stairs elements.
        /// Returns dictionary of run ElementId.Value -> tread length in meters.
        /// </summary>
        private Dictionary<long, double> CalculateAllRunTreadLengths(Document doc)
        {
            var result = new Dictionary<long, double>();

            var stairsElements = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .ToList();

            foreach (var element in stairsElements)
            {
                if (element is not Stairs stairElement) continue;

                try
                {
                    double treadDepth = GetStairTreadDepth(doc, stairElement);
                    if (treadDepth <= 0) continue;

                    var runIds = stairElement.GetStairsRuns();
                    if (runIds == null || runIds.Count == 0) continue;

                    foreach (var runId in runIds)
                    {
                        var runElement = doc.GetElement(runId);
                        if (runElement == null) continue;

                        int riserCount = 0;
                        var riserParam = runElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_NUM_RISERS);
                        if (riserParam != null && riserParam.HasValue)
                            riserCount = riserParam.AsInteger();

                        if (riserCount <= 0) continue;

                        double treadLengthInternal = treadDepth * riserCount;
                        double treadLengthMeters = Math.Round(
                            UnitUtils.ConvertFromInternalUnits(treadLengthInternal, UnitTypeId.Meters), 4);

                        result[runId.Value] = treadLengthMeters;
                    }
                }
                catch
                {
                    // Skip this stair on error
                }
            }

            return result;
        }

        /// <summary>
        /// Generates detailed stair run calculations for manual spot-checking.
        /// </summary>
        private List<StairRunCalculation> GenerateStairRunCalculations(Document doc)
        {
            var calculations = new List<StairRunCalculation>();

            var stairsElements = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Stairs)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent()
                .ToList();

            foreach (var element in stairsElements)
            {
                if (element is not Stairs stairElement) continue;

                try
                {
                    var (treadDepth, treadDepthSource) = GetStairTreadDepthWithSource(doc, stairElement);
                    if (treadDepth <= 0) continue;

                    var runIds = stairElement.GetStairsRuns();
                    if (runIds == null || runIds.Count == 0) continue;

                    foreach (var runId in runIds)
                    {
                        var runElement = doc.GetElement(runId);
                        if (runElement == null) continue;

                        int riserCount = 0;
                        var riserParam = runElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_NUM_RISERS);
                        if (riserParam != null && riserParam.HasValue)
                            riserCount = riserParam.AsInteger();

                        if (riserCount <= 0) continue;

                        double treadDepthMeters = Math.Round(
                            UnitUtils.ConvertFromInternalUnits(treadDepth, UnitTypeId.Meters), 4);
                        double treadLengthMeters = Math.Round(treadDepthMeters * riserCount, 4);

                        calculations.Add(new StairRunCalculation
                        {
                            ElementId = runElement.UniqueId,
                            Name = runElement.Name,
                            TreadDepth = treadDepthMeters,
                            RiserCount = riserCount,
                            CalculatedTreadLength = treadLengthMeters,
                            TreadDepthSource = treadDepthSource
                        });
                    }
                }
                catch { }
            }

            return calculations;
        }

        /// <summary>
        /// Gets tread depth with source information for validation reporting.
        /// </summary>
        private (double treadDepth, string source) GetStairTreadDepthWithSource(Document doc, Stairs stairElement)
        {
            var runIds = stairElement.GetStairsRuns();
            if (runIds != null && runIds.Count > 0)
            {
                var firstRun = doc.GetElement(runIds.First());
                if (firstRun != null)
                {
                    var treadDepthParam = firstRun.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                    if (treadDepthParam != null && treadDepthParam.HasValue && treadDepthParam.StorageType == StorageType.Double)
                    {
                        double val = treadDepthParam.AsDouble();
                        if (val > 0) return (val, "StairRun.STAIRS_ACTUAL_TREAD_DEPTH");
                    }
                }
            }

            var parentTreadParam = stairElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
            if (parentTreadParam != null && parentTreadParam.HasValue && parentTreadParam.StorageType == StorageType.Double)
            {
                double val = parentTreadParam.AsDouble();
                if (val > 0) return (val, "Stairs.STAIRS_ACTUAL_TREAD_DEPTH");
            }

            string[] treadDepthNames = { "Actual Tread Depth", "Tread Depth" };
            foreach (var name in treadDepthNames)
            {
                var param = stairElement.LookupParameter(name);
                if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                {
                    double val = param.AsDouble();
                    if (val > 0) return (val, $"Stairs.LookupParameter(\"{name}\")");
                }
            }

            return (0, "NotFound");
        }

        /// <summary>
        /// Gets tread depth from a Stairs element, trying runs first then parent parameters.
        /// Returns internal units (feet).
        /// </summary>
        private double GetStairTreadDepth(Document doc, Stairs stairElement)
        {
            double treadDepth = 0;

            var runIds = stairElement.GetStairsRuns();
            if (runIds != null && runIds.Count > 0)
            {
                var firstRun = doc.GetElement(runIds.First());
                if (firstRun != null)
                {
                    var treadDepthParam = firstRun.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                    if (treadDepthParam != null && treadDepthParam.HasValue && treadDepthParam.StorageType == StorageType.Double)
                        treadDepth = treadDepthParam.AsDouble();
                }
            }

            if (treadDepth <= 0)
            {
                var parentTreadParam = stairElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                if (parentTreadParam != null && parentTreadParam.HasValue && parentTreadParam.StorageType == StorageType.Double)
                    treadDepth = parentTreadParam.AsDouble();
            }

            if (treadDepth <= 0)
            {
                string[] treadDepthNames = { "Actual Tread Depth", "Tread Depth" };
                foreach (var name in treadDepthNames)
                {
                    var param = stairElement.LookupParameter(name);
                    if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                    {
                        treadDepth = param.AsDouble();
                        if (treadDepth > 0) break;
                    }
                }
            }

            return treadDepth;
        }

        /// <summary>
        /// Calculates horizontal tread length for a stair run element.
        /// Formula: ActualTreadDepth x NumRisers.
        /// Returns meters.
        /// </summary>
        private double CalculateRunTreadLength(Element runElement)
        {
            try
            {
                double treadLengthInternal = 0;

                double treadDepth = 0;
                var treadDepthParam = runElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                if (treadDepthParam != null && treadDepthParam.HasValue && treadDepthParam.StorageType == StorageType.Double)
                    treadDepth = treadDepthParam.AsDouble();

                if (treadDepth <= 0)
                {
                    string[] treadDepthNames = { "Actual Tread Depth", "Tread Depth" };
                    foreach (var name in treadDepthNames)
                    {
                        var param = runElement.LookupParameter(name);
                        if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                        {
                            treadDepth = param.AsDouble();
                            if (treadDepth > 0) break;
                        }
                    }
                }

                int riserCount = 0;
                var riserCountParam = runElement.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_NUM_RISERS);
                if (riserCountParam != null && riserCountParam.HasValue && riserCountParam.StorageType == StorageType.Integer)
                    riserCount = riserCountParam.AsInteger();

                if (riserCount <= 0)
                {
                    string[] riserCountNames = { "Actual Number of Risers", "Number of Risers" };
                    foreach (var name in riserCountNames)
                    {
                        var param = runElement.LookupParameter(name);
                        if (param != null && param.HasValue && param.StorageType == StorageType.Integer)
                        {
                            riserCount = param.AsInteger();
                            if (riserCount > 0) break;
                        }
                    }
                }

                if (treadDepth > 0 && riserCount > 0)
                    treadLengthInternal = treadDepth * riserCount;

                // Fallback: CURVE_ELEM_LENGTH (slope length, but better than 0)
                if (treadLengthInternal <= 0)
                {
                    var lengthParam = runElement.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (lengthParam != null && lengthParam.HasValue)
                        treadLengthInternal = lengthParam.AsDouble();
                }

                // Fallback: "Length" parameter
                if (treadLengthInternal <= 0)
                {
                    var param = runElement.LookupParameter("Length");
                    if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                        treadLengthInternal = param.AsDouble();
                }

                // Fallback: LocationCurve length
                if (treadLengthInternal <= 0 && runElement.Location is LocationCurve locCurve && locCurve.Curve != null)
                {
                    treadLengthInternal = locCurve.Curve.Length;
                }

                if (treadLengthInternal <= 0) return 0;
                return Math.Round(UnitUtils.ConvertFromInternalUnits(treadLengthInternal, UnitTypeId.Meters), 4);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Gets volume from element geometry when parameters don't expose it.
        /// </summary>
        private double GetVolumeFromGeometry(Element element)
        {
            try
            {
                var options = new Options
                {
                    ComputeReferences = false,
                    DetailLevel = ViewDetailLevel.Fine,
                    IncludeNonVisibleObjects = true
                };

                var geomElement = element.get_Geometry(options);
                if (geomElement == null) return 0;

                double volume = SumSolidVolumes(geomElement);
                if (volume <= 0) return 0;
                return UnitUtils.ConvertFromInternalUnits(volume, UnitTypeId.CubicMeters);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Recursively sums solid volumes from geometry, handling nested GeometryInstances.
        /// </summary>
        private double SumSolidVolumes(GeometryElement geomElement)
        {
            double volume = 0;
            foreach (var geomObj in geomElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    volume += solid.Volume;
                }
                else if (geomObj is GeometryInstance geomInstance)
                {
                    var instGeom = geomInstance.GetInstanceGeometry();
                    if (instGeom != null)
                        volume += SumSolidVolumes(instGeom);
                }
            }
            return volume;
        }

        /// <summary>
        /// Gets area from element geometry using top-face analysis.
        /// </summary>
        private double GetAreaFromGeometry(Element element)
        {
            try
            {
                var areaParam = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (areaParam != null && areaParam.HasValue && areaParam.AsDouble() > 0.00001)
                    return UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);

                var options = new Options
                {
                    ComputeReferences = false,
                    DetailLevel = ViewDetailLevel.Fine,
                    IncludeNonVisibleObjects = true
                };

                var geomElement = element.get_Geometry(options);
                if (geomElement != null)
                {
                    double area = GetTopFaceArea(geomElement);
                    if (area > 0)
                        return UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters);
                }

                // Last resort: bounding box footprint
                var bbox = element.get_BoundingBox(null);
                if (bbox != null)
                {
                    double footprint = Math.Abs(bbox.Max.X - bbox.Min.X) * Math.Abs(bbox.Max.Y - bbox.Min.Y);
                    if (footprint > 0)
                        return UnitUtils.ConvertFromInternalUnits(footprint, UnitTypeId.SquareMeters);
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Sums area of top-facing planar faces (normal.Z > 0.9) from geometry.
        /// </summary>
        private double GetTopFaceArea(GeometryElement geomElement)
        {
            double area = 0;
            foreach (var geomObj in geomElement)
            {
                if (geomObj is Solid solid && solid.Faces != null)
                {
                    foreach (Face face in solid.Faces)
                    {
                        if (face is PlanarFace planarFace)
                        {
                            if (Math.Abs(planarFace.FaceNormal.Z) > 0.9)
                                area += face.Area;
                        }
                    }
                }
                else if (geomObj is GeometryInstance geomInstance)
                {
                    var instGeom = geomInstance.GetInstanceGeometry();
                    if (instGeom != null)
                        area += GetTopFaceArea(instGeom);
                }
            }
            return area;
        }

        #endregion Stair Helpers

        #region Host Element Resolution

        /// <summary>
        /// Gets the UniqueId of the host element for hosted elements.
        /// Returns null for non-hosted elements.
        /// </summary>
        private string GetHostElementId(Element e)
        {
            try
            {
                Element host = null;

                if (e is Mullion mullion)
                    host = mullion.Host;
                else if (e is Panel panel)
                    host = panel.Host;
                else if (e is FamilyInstance fi)
                    host = fi.Host;

                if (host != null)
                    return host.UniqueId;
            }
            catch (Exception ex)
            {
                Log.Warning("HostElementId EXCEPTION | {UniqueId} | {Error}", e.UniqueId, ex.Message);
            }

            return null;
        }

        #endregion Host Element Resolution

        #region Parameter Cache Helpers

        private static Dictionary<string, Autodesk.Revit.DB.Parameter> BuildParamCache(Element e)
        {
            var cache = new Dictionary<string, Autodesk.Revit.DB.Parameter>(StringComparer.OrdinalIgnoreCase);
            foreach (Autodesk.Revit.DB.Parameter p in e.Parameters)
            {
                if (p.Definition?.Name != null && !cache.ContainsKey(p.Definition.Name))
                    cache[p.Definition.Name] = p;
            }
            return cache;
        }

        private static Autodesk.Revit.DB.Parameter CacheLookup(Dictionary<string, Autodesk.Revit.DB.Parameter> cache, string name)
        {
            cache.TryGetValue(name, out Autodesk.Revit.DB.Parameter p);
            return p;
        }

        #endregion Parameter Cache Helpers

        #region Core Parameter Helpers

        private ElementType GetElementType(Document doc, Element e)
        {
            ElementId typeId = e.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                return doc.GetElement(typeId) as ElementType;
            }
            return null;
        }

        private string GetElementName(Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            Autodesk.Revit.DB.Parameter mark = e.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
            if (mark != null && mark.HasValue)
            {
                string val = mark.AsString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            Autodesk.Revit.DB.Parameter nameParam = CacheLookup(paramCache, "Name");
            if (nameParam != null && nameParam.HasValue)
            {
                string val = nameParam.AsString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            return null;
        }

        private string GetFunction(Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            Autodesk.Revit.DB.Parameter funcParam = e.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT);
            if (funcParam != null && funcParam.HasValue)
            {
                return funcParam.AsInteger() == 1 ? "Structural" : "Non-Structural";
            }

            Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, "Function");
            if (p != null && p.HasValue)
            {
                return p.AsValueString() ?? p.AsString();
            }

            Autodesk.Revit.DB.Parameter sfParam = e.get_Parameter(BuiltInParameter.INSTANCE_STRUCT_USAGE_PARAM);
            if (sfParam != null && sfParam.HasValue)
            {
                return sfParam.AsValueString();
            }

            return null;
        }

        private string GetAssemblyDescription(Document doc, Element e, ElementType type)
        {
            Autodesk.Revit.DB.Parameter instParam = e.get_Parameter(BuiltInParameter.UNIFORMAT_DESCRIPTION);
            if (instParam != null && instParam.HasValue)
            {
                string val = instParam.AsString() ?? instParam.AsValueString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            if (type != null)
            {
                Autodesk.Revit.DB.Parameter typeParam = type.get_Parameter(BuiltInParameter.UNIFORMAT_DESCRIPTION);
                if (typeParam != null && typeParam.HasValue)
                {
                    string val = typeParam.AsString() ?? typeParam.AsValueString();
                    if (!string.IsNullOrWhiteSpace(val))
                        return val;
                }
            }

            return null;
        }

        private string GetTypeModelParameter(ElementType type)
        {
            if (type == null) return null;

            Autodesk.Revit.DB.Parameter p = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL);
            if (p != null && p.HasValue)
            {
                string val = p.AsString() ?? p.AsValueString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            return null;
        }

        #endregion Core Parameter Helpers

        #region Level Extraction

        private string GetLevelRobust(Document doc, Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            string[] levelParamNames = new[]
            {
                "Level", "LevelName", "Reference Level", "Schedule Level",
                "Base Level", "Base Constraint", "Constraint Level", "Host Level"
            };

            foreach (string paramName in levelParamNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, paramName);
                if (p != null && p.HasValue)
                {
                    string value = GetParameterDisplayValue(doc, p);
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }

            BuiltInParameter[] bipFallbacks = new[]
            {
                BuiltInParameter.FAMILY_LEVEL_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                BuiltInParameter.WALL_BASE_CONSTRAINT,
                BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                BuiltInParameter.STAIRS_BASE_LEVEL_PARAM,
                BuiltInParameter.ASSOCIATED_LEVEL,
                BuiltInParameter.LEVEL_PARAM
            };

            foreach (var bip in bipFallbacks)
            {
                Autodesk.Revit.DB.Parameter p = e.get_Parameter(bip);
                if (p != null && p.HasValue)
                {
                    string value = GetParameterDisplayValue(doc, p);
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }

            if (e.LevelId != ElementId.InvalidElementId)
            {
                Element lvl = doc.GetElement(e.LevelId);
                if (lvl != null && !string.IsNullOrWhiteSpace(lvl.Name))
                    return lvl.Name;
            }

            return "Unassigned";
        }

        private double? GetLevelElevation(Document doc, Element e, string levelName, Dictionary<string, Level> levelCache)
        {
            if (string.IsNullOrEmpty(levelName) || levelName == "Unassigned")
                return null;

            if (levelCache.TryGetValue(levelName, out Level level))
            {
                return UnitUtils.ConvertFromInternalUnits(level.Elevation, UnitTypeId.Meters);
            }

            if (e.LevelId != ElementId.InvalidElementId)
            {
                Level directLevel = doc.GetElement(e.LevelId) as Level;
                if (directLevel != null)
                {
                    return UnitUtils.ConvertFromInternalUnits(directLevel.Elevation, UnitTypeId.Meters);
                }
            }

            return null;
        }

        private string GetParameterDisplayValue(Document doc, Autodesk.Revit.DB.Parameter p)
        {
            if (p == null || !p.HasValue) return null;

            if (p.StorageType == StorageType.ElementId)
            {
                ElementId id = p.AsElementId();
                if (id != ElementId.InvalidElementId)
                {
                    Element refEl = doc.GetElement(id);
                    return refEl?.Name;
                }
                return null;
            }

            return p.AsValueString() ?? p.AsString();
        }

        #endregion Level Extraction

        #region Phase Extraction

        private string GetPhaseRobust(Document doc, Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            Autodesk.Revit.DB.Parameter phaseCreated = e.get_Parameter(BuiltInParameter.PHASE_CREATED);
            if (phaseCreated != null && phaseCreated.HasValue)
            {
                string val = GetParameterDisplayValue(doc, phaseCreated);
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            Autodesk.Revit.DB.Parameter phase = CacheLookup(paramCache, "Phase");
            if (phase != null && phase.HasValue)
            {
                string val = GetParameterDisplayValue(doc, phase);
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            return "Existing";
        }

        #endregion Phase Extraction

        #region Quantity Extraction

        private double GetVolumeRobust(Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            double rawValue = 0;

            string[] paramNames = { "Volume", "Net Volume" };
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (IsValidDouble(p))
                {
                    rawValue = p.AsDouble();
                    break;
                }
            }

            if (rawValue == 0)
            {
                Autodesk.Revit.DB.Parameter p = e.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (IsValidDouble(p))
                    rawValue = p.AsDouble();
            }

            if (rawValue <= 0) return 0;
            return Math.Round(UnitUtils.ConvertFromInternalUnits(rawValue, UnitTypeId.CubicMeters), 4);
        }

        private double GetAreaRobust(Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            double rawValue = 0;

            string[] paramNames = { "Area", "Net Area", "Surface Area" };
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (IsValidDouble(p))
                {
                    rawValue = p.AsDouble();
                    break;
                }
            }

            if (rawValue == 0)
            {
                Autodesk.Revit.DB.Parameter p = e.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (IsValidDouble(p))
                    rawValue = p.AsDouble();
            }

            if (rawValue <= 0) return 0;
            return Math.Round(UnitUtils.ConvertFromInternalUnits(rawValue, UnitTypeId.SquareMeters), 4);
        }

        /// <summary>
        /// Gets length in meters with category-aware fallback chain.
        /// Walls use height, columns use instance length, floors/ceilings/roofs use perimeter.
        /// </summary>
        private double GetLengthRobust(Element e, Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache)
        {
            double rawValue = 0;
            long catId = e.Category?.Id.Value ?? 0;

            // Category-specific: walls -- height is their primary linear measurement
            if (catId == (long)BuiltInCategory.OST_Walls)
            {
                var wallHeight = e.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (IsValidDouble(wallHeight))
                    return Math.Round(UnitUtils.ConvertFromInternalUnits(wallHeight.AsDouble(), UnitTypeId.Meters), 4);
            }

            // Category-specific: columns -- height from level calc or instance length
            if (catId == (long)BuiltInCategory.OST_StructuralColumns)
            {
                var colLength = e.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM)
                             ?? e.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
                if (colLength != null && IsValidDouble(colLength))
                    return Math.Round(UnitUtils.ConvertFromInternalUnits(colLength.AsDouble(), UnitTypeId.Meters), 4);
            }

            // Category-specific: floors/ceilings/roofs -- perimeter
            if (catId == (long)BuiltInCategory.OST_Floors ||
                catId == (long)BuiltInCategory.OST_Ceilings ||
                catId == (long)BuiltInCategory.OST_Roofs)
            {
                Autodesk.Revit.DB.Parameter perimParam = CacheLookup(paramCache, "Perimeter")
                                    ?? e.get_Parameter(BuiltInParameter.HOST_PERIMETER_COMPUTED);
                if (IsValidDouble(perimParam))
                    return Math.Round(UnitUtils.ConvertFromInternalUnits(perimParam.AsDouble(), UnitTypeId.Meters), 4);
            }

            // Generic: true length parameters only
            string[] paramNames = new[] { "Length", "Cut Length" };

            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (IsValidDouble(p))
                {
                    rawValue = p.AsDouble();
                    break;
                }
            }

            if (rawValue == 0)
            {
                BuiltInParameter[] bips = new[]
                {
                    BuiltInParameter.CURVE_ELEM_LENGTH,
                    BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH,
                    BuiltInParameter.INSTANCE_LENGTH_PARAM
                };

                foreach (var bip in bips)
                {
                    Autodesk.Revit.DB.Parameter p = e.get_Parameter(bip);
                    if (IsValidDouble(p))
                    {
                        rawValue = p.AsDouble();
                        break;
                    }
                }
            }

            if (rawValue <= 0) return 0;
            return Math.Round(UnitUtils.ConvertFromInternalUnits(rawValue, UnitTypeId.Meters), 4);
        }

        private bool IsValidDouble(Autodesk.Revit.DB.Parameter p)
        {
            return p != null && p.HasValue && p.StorageType == StorageType.Double && p.AsDouble() > 0.00001;
        }

        #endregion Quantity Extraction

        #region Type Parameter Extraction

        private string GetTypeParameter(Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, ElementType type, string paramName, BuiltInParameter bipFallback)
        {
            if (type == null) return null;

            Autodesk.Revit.DB.Parameter p = typeParamCache != null ? CacheLookup(typeParamCache, paramName) : null;
            if (p != null && p.HasValue)
            {
                string val = p.AsString() ?? p.AsValueString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            if (bipFallback != BuiltInParameter.INVALID)
            {
                Autodesk.Revit.DB.Parameter bipParam = type.get_Parameter(bipFallback);
                if (bipParam != null && bipParam.HasValue)
                {
                    string val = bipParam.AsString() ?? bipParam.AsValueString();
                    if (!string.IsNullOrWhiteSpace(val))
                        return val;
                }
            }

            return null;
        }

        private string GetAssemblyCode(Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, ElementType type)
        {
            if (type == null) return null;

            Autodesk.Revit.DB.Parameter p = typeParamCache != null ? CacheLookup(typeParamCache, "Assembly Code") : null;
            if (p != null && p.HasValue)
            {
                string val = p.AsString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            Autodesk.Revit.DB.Parameter bip = type.get_Parameter(BuiltInParameter.UNIFORMAT_CODE);
            if (bip != null && bip.HasValue)
            {
                string val = bip.AsString();
                if (!string.IsNullOrWhiteSpace(val))
                    return val;
            }

            return null;
        }

        #endregion Type Parameter Extraction

        #region BI_ Commercial Parameter Extraction

        private string GetBIParameter(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, params string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    string val = p.AsString() ?? p.AsValueString();
                    if (!string.IsNullOrWhiteSpace(val))
                        return val;
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        string val = p.AsString() ?? p.AsValueString();
                        if (!string.IsNullOrWhiteSpace(val))
                            return val;
                    }
                }
            }

            return null;
        }

        private bool? GetBIBoolParameter(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, params string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    if (p.StorageType == StorageType.Integer)
                        return p.AsInteger() == 1;
                    else if (p.StorageType == StorageType.String)
                    {
                        string val = p.AsString()?.Trim().ToLower();
                        if (!string.IsNullOrEmpty(val))
                            return val == "true" || val == "yes" || val == "1";
                    }
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        if (p.StorageType == StorageType.Integer)
                            return p.AsInteger() == 1;
                        else if (p.StorageType == StorageType.String)
                        {
                            string val = p.AsString()?.Trim().ToLower();
                            if (!string.IsNullOrEmpty(val))
                                return val == "true" || val == "yes" || val == "1";
                        }
                    }
                }
            }

            return null;
        }

        private double? GetBIDoubleParameter(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, params string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    if (p.StorageType == StorageType.Double)
                    {
                        double val = p.AsDouble();
                        if (Math.Abs(val) > 0.00001)
                            return val;
                    }
                    else if (p.StorageType == StorageType.Integer)
                    {
                        int val = p.AsInteger();
                        if (val != 0)
                            return val;
                    }
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        if (p.StorageType == StorageType.Double)
                        {
                            double val = p.AsDouble();
                            if (Math.Abs(val) > 0.00001)
                                return val;
                        }
                        else if (p.StorageType == StorageType.Integer)
                        {
                            int val = p.AsInteger();
                            if (val != 0)
                                return val;
                        }
                    }
                }
            }

            return null;
        }

        #endregion BI_ Commercial Parameter Extraction

        #region Legacy BOQ/Contract Parameter Extraction

        private string GetBOQParameter(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache, params string[] paramNames)
        {
            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    string val = p.AsString() ?? p.AsValueString();
                    if (!string.IsNullOrWhiteSpace(val))
                        return val;
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        string val = p.AsString() ?? p.AsValueString();
                        if (!string.IsNullOrWhiteSpace(val))
                            return val;
                    }
                }
            }

            return null;
        }

        private bool? GetIsMedida(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache)
        {
            string[] paramNames = { "IsMedida", "Is Medida" };

            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    if (p.StorageType == StorageType.Integer)
                        return p.AsInteger() == 1;
                    else if (p.StorageType == StorageType.String)
                    {
                        string val = p.AsString()?.Trim().ToLower();
                        if (!string.IsNullOrEmpty(val))
                            return val == "true" || val == "yes" || val == "1";
                    }
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        if (p.StorageType == StorageType.Integer)
                            return p.AsInteger() == 1;
                        else if (p.StorageType == StorageType.String)
                        {
                            string val = p.AsString()?.Trim().ToLower();
                            if (!string.IsNullOrEmpty(val))
                                return val == "true" || val == "yes" || val == "1";
                        }
                    }
                }
            }

            return null;
        }

        private double? GetTotalContractPrice(Dictionary<string, Autodesk.Revit.DB.Parameter> paramCache, Dictionary<string, Autodesk.Revit.DB.Parameter> typeParamCache)
        {
            string[] paramNames = { "TotalContractPrice", "Total Contract Price", "TotalPrice", "Total Price" };

            foreach (string name in paramNames)
            {
                Autodesk.Revit.DB.Parameter p = CacheLookup(paramCache, name);
                if (p != null && p.HasValue)
                {
                    if (p.StorageType == StorageType.Double)
                    {
                        double val = p.AsDouble();
                        if (val > 0) return val;
                    }
                    else if (p.StorageType == StorageType.Integer)
                    {
                        int val = p.AsInteger();
                        if (val > 0) return val;
                    }
                }
            }

            if (typeParamCache != null)
            {
                foreach (string name in paramNames)
                {
                    Autodesk.Revit.DB.Parameter p = CacheLookup(typeParamCache, name);
                    if (p != null && p.HasValue)
                    {
                        if (p.StorageType == StorageType.Double)
                        {
                            double val = p.AsDouble();
                            if (val > 0) return val;
                        }
                        else if (p.StorageType == StorageType.Integer)
                        {
                            int val = p.AsInteger();
                            if (val > 0) return val;
                        }
                    }
                }
            }

            return null;
        }

        #endregion Legacy BOQ/Contract Parameter Extraction
    }
}
