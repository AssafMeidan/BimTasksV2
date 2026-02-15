using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Creates validation schedules in Revit for cross-checking exported data.
    /// Schedules show count, volume, area, and length totals by category.
    /// Migrated from OldApp using Revit 2025 APIs (ElementId.Value instead of IntegerValue).
    /// </summary>
    public class ValidationScheduleService : IValidationScheduleService
    {
        private const string SchedulePrefix = "BIM_Validation_";

        /// <summary>
        /// Categories that need validation schedules.
        /// Excludes Stair Runs and Landings (calculated values, not schedulable).
        /// </summary>
        private static readonly Dictionary<string, BuiltInCategory> SchedulableCategories = new Dictionary<string, BuiltInCategory>
        {
            { "Walls", BuiltInCategory.OST_Walls },
            { "Floors", BuiltInCategory.OST_Floors },
            { "Ceilings", BuiltInCategory.OST_Ceilings },
            { "Roofs", BuiltInCategory.OST_Roofs },
            { "Doors", BuiltInCategory.OST_Doors },
            { "Windows", BuiltInCategory.OST_Windows },
            { "StructuralColumns", BuiltInCategory.OST_StructuralColumns },
            { "StructuralFraming", BuiltInCategory.OST_StructuralFraming },
            { "StructuralFoundation", BuiltInCategory.OST_StructuralFoundation },
            { "Stairs", BuiltInCategory.OST_Stairs },
            { "Railings", BuiltInCategory.OST_Railings },
            { "Furniture", BuiltInCategory.OST_Furniture },
            { "GenericModels", BuiltInCategory.OST_GenericModel },
            { "Casework", BuiltInCategory.OST_Casework },
            { "MechanicalEquipment", BuiltInCategory.OST_MechanicalEquipment },
            { "PlumbingFixtures", BuiltInCategory.OST_PlumbingFixtures },
            { "ElectricalFixtures", BuiltInCategory.OST_ElectricalFixtures },
            { "ElectricalEquipment", BuiltInCategory.OST_ElectricalEquipment },
            { "LightingFixtures", BuiltInCategory.OST_LightingFixtures },
            { "Ducts", BuiltInCategory.OST_DuctCurves },
            { "Pipes", BuiltInCategory.OST_PipeCurves },
            { "CableTrays", BuiltInCategory.OST_CableTray },
            { "Conduits", BuiltInCategory.OST_Conduit }
        };

        #region Public Methods

        /// <inheritdoc/>
        public List<string> EnsureSchedules(Document doc)
        {
            var createdSchedules = new List<string>();

            using (Transaction trans = new Transaction(doc, "Create Validation Schedules"))
            {
                trans.Start();

                foreach (var kvp in SchedulableCategories)
                {
                    string scheduleName = SchedulePrefix + kvp.Key;

                    // Check if schedule already exists
                    var existingSchedule = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSchedule))
                        .Cast<ViewSchedule>()
                        .FirstOrDefault(vs => vs.Name == scheduleName);

                    if (existingSchedule == null)
                    {
                        try
                        {
                            var schedule = CreateValidationSchedule(doc, scheduleName, kvp.Value);
                            if (schedule != null)
                            {
                                createdSchedules.Add(scheduleName);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Debug("Could not create schedule {Name}: {Error}", scheduleName, ex.Message);
                        }
                    }
                }

                trans.Commit();
            }

            return createdSchedules;
        }

        /// <inheritdoc/>
        public Dictionary<string, ValidationScheduleTotals> GetTotals(Document doc, HashSet<long> validCreationPhaseIds = null)
        {
            var result = new Dictionary<string, ValidationScheduleTotals>();

            var schedules = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => vs.Name.StartsWith(SchedulePrefix))
                .ToList();

            foreach (var schedule in schedules)
            {
                try
                {
                    string categoryName = schedule.Name.Substring(SchedulePrefix.Length);
                    var totals = ExtractScheduleTotals(doc, schedule, validCreationPhaseIds);
                    result[categoryName] = totals;
                }
                catch (Exception ex)
                {
                    Log.Debug("Error extracting totals from {Schedule}: {Error}", schedule.Name, ex.Message);
                }
            }

            return result;
        }

        /// <inheritdoc/>
        public Dictionary<string, ValidationScheduleTotals> GetFamilyTotals(Document doc, string scheduleName, HashSet<long> validCreationPhaseIds = null)
        {
            var result = new Dictionary<string, ValidationScheduleTotals>();

            var schedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(vs => vs.Name == SchedulePrefix + scheduleName);

            if (schedule == null)
                return result;

            var elements = new FilteredElementCollector(doc, schedule.Id)
                .WhereElementIsNotElementType()
                .ToList();

            return CalculateFamilyTotals(doc, elements, validCreationPhaseIds);
        }

        #endregion Public Methods

        #region Schedule Creation

        /// <summary>
        /// Creates a validation schedule for a specific category.
        /// Shows: Family and Type grouped with subtotals, plus grand totals.
        /// </summary>
        private ViewSchedule CreateValidationSchedule(Document doc, string scheduleName, BuiltInCategory category)
        {
            ElementId categoryId = new ElementId(category);
            ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, categoryId);
            schedule.Name = scheduleName;

            // Get schedulable fields
            var schedulableFields = schedule.Definition.GetSchedulableFields();
            var fieldDict = new Dictionary<string, SchedulableField>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in schedulableFields)
            {
                string name = f.GetName(doc).ToLower();
                if (!fieldDict.ContainsKey(name))
                    fieldDict[name] = f;
            }

            // Add Family and Type field for grouping
            var familyTypeField = AddFieldIfExists(doc, schedule, fieldDict, "family and type", "family", "type name", "type");

            // Add Count field (calculated)
            try
            {
                schedule.Definition.AddField(ScheduleFieldType.Count);
            }
            catch { }

            // Add Volume field if available
            AddFieldIfExists(doc, schedule, fieldDict, "volume");

            // Add Area field if available
            AddFieldIfExists(doc, schedule, fieldDict, "area");

            // Add Length field if available
            AddFieldIfExists(doc, schedule, fieldDict, "length", "cut length");

            // Sort by Family and Type
            if (familyTypeField != null)
            {
                var sortGroup = new ScheduleSortGroupField(familyTypeField.FieldId);
                sortGroup.ShowHeader = true;
                sortGroup.ShowFooter = true;
                sortGroup.ShowBlankLine = false;
                schedule.Definition.AddSortGroupField(sortGroup);
            }

            // Configure schedule: show items with grand totals
            schedule.Definition.IsItemized = true;
            schedule.Definition.ShowGrandTotal = true;
            schedule.Definition.ShowGrandTotalTitle = true;
            schedule.Definition.ShowGrandTotalCount = true;

            return schedule;
        }

        /// <summary>
        /// Adds a field to the schedule if it exists, trying multiple field names.
        /// </summary>
        private ScheduleField AddFieldIfExists(
            Document doc,
            ViewSchedule schedule,
            Dictionary<string, SchedulableField> fieldDict,
            params string[] fieldNames)
        {
            foreach (var name in fieldNames)
            {
                if (fieldDict.TryGetValue(name.ToLower(), out var schedulableField))
                {
                    try
                    {
                        return schedule.Definition.AddField(schedulableField);
                    }
                    catch
                    {
                        // Field might already exist or can't be added
                    }
                }
            }
            return null;
        }

        #endregion Schedule Creation

        #region Totals Extraction

        /// <summary>
        /// Extracts totals from a validation schedule by querying elements directly.
        /// </summary>
        private ValidationScheduleTotals ExtractScheduleTotals(Document doc, ViewSchedule schedule, HashSet<long> validCreationPhaseIds = null)
        {
            var totals = new ValidationScheduleTotals();

            var elements = new FilteredElementCollector(doc, schedule.Id)
                .WhereElementIsNotElementType()
                .ToList();

            foreach (var element in elements)
            {
                // Phase filter: skip future or demolished elements
                if (!PassesPhaseFilter(element, validCreationPhaseIds))
                    continue;

                // Get geometry values
                double vol = 0, area = 0, len = 0;

                var volumeParam = element.LookupParameter("Volume") ?? element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (volumeParam != null && volumeParam.HasValue && volumeParam.StorageType == StorageType.Double)
                    vol = UnitUtils.ConvertFromInternalUnits(volumeParam.AsDouble(), UnitTypeId.CubicMeters);

                var areaParam = element.LookupParameter("Area") ?? element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (areaParam != null && areaParam.HasValue && areaParam.StorageType == StorageType.Double)
                    area = UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);

                len = GetLengthCategoryAware(element);

                // Skip zero-geometry elements (consistent with JSON export filter)
                if (vol <= 0 && area <= 0 && len <= 0)
                    continue;

                totals.Count++;
                totals.Volume += vol;
                totals.Area += area;
                totals.Length += len;
            }

            totals.Volume = Math.Round(totals.Volume, 4);
            totals.Area = Math.Round(totals.Area, 4);
            totals.Length = Math.Round(totals.Length, 4);

            return totals;
        }

        /// <summary>
        /// Calculates family-level totals from a list of elements.
        /// </summary>
        private Dictionary<string, ValidationScheduleTotals> CalculateFamilyTotals(Document doc, List<Element> elements, HashSet<long> validCreationPhaseIds = null)
        {
            var result = new Dictionary<string, ValidationScheduleTotals>();

            foreach (var element in elements)
            {
                if (!PassesPhaseFilter(element, validCreationPhaseIds))
                    continue;

                double vol = 0, area = 0, len = 0;

                var volumeParam = element.LookupParameter("Volume") ?? element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (volumeParam != null && volumeParam.HasValue && volumeParam.StorageType == StorageType.Double)
                    vol = UnitUtils.ConvertFromInternalUnits(volumeParam.AsDouble(), UnitTypeId.CubicMeters);

                var areaParam = element.LookupParameter("Area") ?? element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (areaParam != null && areaParam.HasValue && areaParam.StorageType == StorageType.Double)
                    area = UnitUtils.ConvertFromInternalUnits(areaParam.AsDouble(), UnitTypeId.SquareMeters);

                len = GetLengthCategoryAware(element);

                if (vol <= 0 && area <= 0 && len <= 0)
                    continue;

                // Get family name
                string familyName = "";
                var typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    var elementType = doc.GetElement(typeId) as ElementType;
                    familyName = elementType?.FamilyName ?? "";
                }
                if (string.IsNullOrEmpty(familyName))
                    familyName = "(No Family)";

                if (!result.TryGetValue(familyName, out var totals))
                {
                    totals = new ValidationScheduleTotals();
                    result[familyName] = totals;
                }

                totals.Count++;
                totals.Volume += vol;
                totals.Area += area;
                totals.Length += len;
            }

            // Round all values
            foreach (var kvp in result)
            {
                kvp.Value.Volume = Math.Round(kvp.Value.Volume, 4);
                kvp.Value.Area = Math.Round(kvp.Value.Area, 4);
                kvp.Value.Length = Math.Round(kvp.Value.Length, 4);
            }

            return result;
        }

        #endregion Totals Extraction

        #region Helpers

        /// <summary>
        /// Checks whether an element passes the phase filter (Show Complete logic).
        /// Uses ElementId.Value for Revit 2025.
        /// </summary>
        private bool PassesPhaseFilter(Element element, HashSet<long> validCreationPhaseIds)
        {
            if (validCreationPhaseIds == null)
                return true;

            var createdParam = element.get_Parameter(BuiltInParameter.PHASE_CREATED);
            var demolishedParam = element.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);

            if (createdParam != null && createdParam.HasValue)
            {
                long createdId = createdParam.AsElementId().Value;
                if (!validCreationPhaseIds.Contains(createdId))
                    return false;
            }

            if (demolishedParam != null && demolishedParam.HasValue)
            {
                long demolishedId = demolishedParam.AsElementId().Value;
                if (demolishedId != -1 && validCreationPhaseIds.Contains(demolishedId))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Gets length using the same category-aware logic as DataExtractorService.GetLengthRobust.
        /// Walls use height, columns use instance length, floors/ceilings/roofs use perimeter.
        /// Uses ElementId.Value for Revit 2025.
        /// </summary>
        private double GetLengthCategoryAware(Element element)
        {
            long catId = element.Category?.Id.Value ?? 0;

            // Walls: height is their primary linear measurement
            if (catId == (long)BuiltInCategory.OST_Walls)
            {
                var wallHeight = element.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (wallHeight != null && wallHeight.HasValue && wallHeight.StorageType == StorageType.Double)
                    return UnitUtils.ConvertFromInternalUnits(wallHeight.AsDouble(), UnitTypeId.Meters);
            }

            // Columns: instance length or family height
            if (catId == (long)BuiltInCategory.OST_StructuralColumns)
            {
                var colLength = element.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM)
                             ?? element.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
                if (colLength != null && colLength.HasValue && colLength.StorageType == StorageType.Double)
                    return UnitUtils.ConvertFromInternalUnits(colLength.AsDouble(), UnitTypeId.Meters);
            }

            // Floors/Ceilings/Roofs: perimeter
            if (catId == (long)BuiltInCategory.OST_Floors ||
                catId == (long)BuiltInCategory.OST_Ceilings ||
                catId == (long)BuiltInCategory.OST_Roofs)
            {
                var perimParam = element.LookupParameter("Perimeter")
                              ?? element.get_Parameter(BuiltInParameter.HOST_PERIMETER_COMPUTED);
                if (perimParam != null && perimParam.HasValue && perimParam.StorageType == StorageType.Double)
                    return UnitUtils.ConvertFromInternalUnits(perimParam.AsDouble(), UnitTypeId.Meters);
            }

            // Generic: Length or CURVE_ELEM_LENGTH
            var lengthParam = element.LookupParameter("Length") ?? element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
            if (lengthParam != null && lengthParam.HasValue && lengthParam.StorageType == StorageType.Double)
                return UnitUtils.ConvertFromInternalUnits(lengthParam.AsDouble(), UnitTypeId.Meters);

            return 0;
        }

        #endregion Helpers
    }
}
