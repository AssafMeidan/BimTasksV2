using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Implementation of the element calculation service.
    /// Uses Revit 2025 API (ElementId.Value instead of IntegerValue).
    /// </summary>
    public class ElementCalculationService : IElementCalculationService
    {
        private const string ServiceName = nameof(ElementCalculationService);

        public ElementCalculationResult CalculateElements(Document doc, ICollection<ElementId> selectedIds)
        {
            var stopwatch = Stopwatch.StartNew();
            var result = new ElementCalculationResult();

            try
            {
                result.SelectedElementCount = selectedIds.Count;
                Log.Information("[{Service}] Starting calculation for {Count} selected elements",
                    ServiceName, selectedIds.Count);

                if (selectedIds.Count == 0)
                {
                    result.HasErrors = true;
                    result.Errors.Add("No elements selected");
                    return result;
                }

                LogSelectionBreakdown(doc, selectedIds, result);

                CalculateWalls(doc, selectedIds, result);
                CalculateFloors(doc, selectedIds, result);
                CalculateStructuralFraming(doc, selectedIds, result);
                CalculateStructuralColumns(doc, selectedIds, result);
                CalculateFoundations(doc, selectedIds, result);
                CalculateStairs(doc, selectedIds, result);
                CalculateRailings(doc, selectedIds, result);
                CalculateDoors(doc, selectedIds, result);
                CalculateWindows(doc, selectedIds, result);
                CalculateMaterialVolumes(doc, selectedIds, result);

                stopwatch.Stop();
                result.CalculationTimeMs = stopwatch.ElapsedMilliseconds;

                Log.Information("[{Service}] Calculation complete in {Time}ms. Total concrete: {Volume:F2} m3",
                    ServiceName, result.CalculationTimeMs, result.TotalConcreteVolumeM3);

                return result;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                result.CalculationTimeMs = stopwatch.ElapsedMilliseconds;
                result.HasErrors = true;
                result.Errors.Add(ex.Message);
                Log.Error(ex, "[{Service}] Calculation failed", ServiceName);
                return result;
            }
        }

        #region Selection Analysis

        private void LogSelectionBreakdown(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            var categoryCount = new Dictionary<string, int>();

            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element == null) continue;

                string catName = element.Category?.Name ?? "No Category";

                if (!categoryCount.ContainsKey(catName))
                    categoryCount[catName] = 0;
                categoryCount[catName]++;
            }

            foreach (var kvp in categoryCount.OrderByDescending(x => x.Value))
            {
                result.CategoryBreakdown[kvp.Key] = kvp.Value;
            }

            Log.Debug("[{Service}] Selection breakdown: {Categories}",
                ServiceName, string.Join(", ", categoryCount.Select(kv => $"{kv.Key}={kv.Value}")));
        }

        #endregion

        #region Walls

        private void CalculateWalls(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                var walls = new FilteredElementCollector(doc, selectedIds)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();

                double totalArea = 0;
                double totalVolume = 0;

                foreach (var wall in walls)
                {
                    try
                    {
                        var areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                        var volumeParam = wall.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);

                        if (areaParam != null && areaParam.HasValue)
                            totalArea += areaParam.AsDouble();
                        if (volumeParam != null && volumeParam.HasValue)
                            totalVolume += volumeParam.AsDouble();

                        result.WallCount++;
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("[{Service}] Error calculating wall {Id}: {Error}",
                            ServiceName, wall.Id.Value, ex.Message);
                    }
                }

                result.WallAreaM2 = ConvertToSquareMeters(totalArea);
                result.WallVolumeM3 = ConvertToCubicMeters(totalVolume);

                Log.Information("[{Service}] Walls: {Count}, Area: {Area:F2} m2, Volume: {Vol:F2} m3",
                    ServiceName, result.WallCount, result.WallAreaM2, result.WallVolumeM3);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in wall calculation", ServiceName);
                result.Errors.Add($"Wall calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Floors

        private void CalculateFloors(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                var floors = new FilteredElementCollector(doc, selectedIds)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .Cast<Floor>()
                    .ToList();

                double totalArea = 0;
                double totalVolume = 0;

                foreach (var floor in floors)
                {
                    try
                    {
                        var areaParam = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                        var volumeParam = floor.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);

                        if (areaParam != null && areaParam.HasValue)
                            totalArea += areaParam.AsDouble();
                        if (volumeParam != null && volumeParam.HasValue)
                            totalVolume += volumeParam.AsDouble();

                        result.FloorCount++;
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("[{Service}] Error calculating floor {Id}: {Error}",
                            ServiceName, floor.Id.Value, ex.Message);
                    }
                }

                result.FloorAreaM2 = ConvertToSquareMeters(totalArea);
                result.FloorVolumeM3 = ConvertToCubicMeters(totalVolume);

                Log.Information("[{Service}] Floors: {Count}, Area: {Area:F2} m2, Volume: {Vol:F2} m3",
                    ServiceName, result.FloorCount, result.FloorAreaM2, result.FloorVolumeM3);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in floor calculation", ServiceName);
                result.Errors.Add($"Floor calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Structural Framing (Beams)

        private void CalculateStructuralFraming(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                double totalVolume = 0;
                double totalLength = 0;
                int beamCount = 0;

                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category == null) continue;

                    if (element.Category.Id.Value == (long)BuiltInCategory.OST_StructuralFraming)
                    {
                        beamCount++;

                        // Volume
                        var volumeParam = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                        if (volumeParam != null && volumeParam.HasValue)
                            totalVolume += volumeParam.AsDouble();

                        // Length
                        double length = 0;

                        if (element.Location is LocationCurve locCurve && locCurve.Curve != null)
                        {
                            length = locCurve.Curve.Length;
                        }

                        if (length <= 0)
                        {
                            var param = element.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM)
                                     ?? element.get_Parameter(BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH)
                                     ?? element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                            if (param != null && param.HasValue)
                                length = param.AsDouble();
                        }

                        if (length <= 0)
                        {
                            var bbox = element.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                double dx = Math.Abs(bbox.Max.X - bbox.Min.X);
                                double dy = Math.Abs(bbox.Max.Y - bbox.Min.Y);
                                length = Math.Max(dx, dy);
                            }
                        }

                        totalLength += length;
                    }
                }

                result.BeamCount = beamCount;
                result.BeamVolumeM3 = ConvertToCubicMeters(totalVolume);
                result.BeamLengthM = ConvertToMeters(totalLength);

                Log.Information("[{Service}] Beams: {Count}, Volume: {Vol:F2} m3, Length: {Len:F2} m",
                    ServiceName, result.BeamCount, result.BeamVolumeM3, result.BeamLengthM);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in beam calculation", ServiceName);
                result.Errors.Add($"Beam calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Structural Columns

        private void CalculateStructuralColumns(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                double totalVolume = 0;
                double totalHeight = 0;
                int columnCount = 0;

                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category == null) continue;

                    if (element.Category.Id.Value == (long)BuiltInCategory.OST_StructuralColumns)
                    {
                        columnCount++;

                        // Volume
                        var volumeParam = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                        if (volumeParam != null && volumeParam.HasValue)
                            totalVolume += volumeParam.AsDouble();

                        // Height from levels
                        double height = 0;
                        try
                        {
                            var baseLevelParam = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                            var topLevelParam = element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);

                            if (baseLevelParam != null && topLevelParam != null &&
                                baseLevelParam.AsElementId() != ElementId.InvalidElementId &&
                                topLevelParam.AsElementId() != ElementId.InvalidElementId)
                            {
                                var baseLevel = doc.GetElement(baseLevelParam.AsElementId()) as Level;
                                var topLevel = doc.GetElement(topLevelParam.AsElementId()) as Level;

                                if (baseLevel != null && topLevel != null)
                                {
                                    double baseElev = baseLevel.Elevation;
                                    double topElev = topLevel.Elevation;

                                    var baseOffsetParam = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                                    var topOffsetParam = element.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);

                                    if (baseOffsetParam != null && baseOffsetParam.HasValue)
                                        baseElev += baseOffsetParam.AsDouble();
                                    if (topOffsetParam != null && topOffsetParam.HasValue)
                                        topElev += topOffsetParam.AsDouble();

                                    height = Math.Abs(topElev - baseElev);
                                }
                            }
                        }
                        catch { /* fallback below */ }

                        if (height <= 0 && element.Location is LocationCurve locCurve && locCurve.Curve != null)
                            height = locCurve.Curve.Length;

                        if (height <= 0)
                        {
                            var param = element.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM)
                                     ?? element.get_Parameter(BuiltInParameter.FAMILY_HEIGHT_PARAM);
                            if (param != null && param.HasValue)
                                height = Math.Abs(param.AsDouble());
                        }

                        if (height <= 0)
                        {
                            var bbox = element.get_BoundingBox(null);
                            if (bbox != null)
                                height = Math.Abs(bbox.Max.Z - bbox.Min.Z);
                        }

                        totalHeight += height;
                    }
                }

                result.ColumnCount = columnCount;
                result.ColumnVolumeM3 = ConvertToCubicMeters(totalVolume);
                result.ColumnHeightM = ConvertToMeters(totalHeight);

                Log.Information("[{Service}] Columns: {Count}, Volume: {Vol:F2} m3, Height: {H:F2} m",
                    ServiceName, result.ColumnCount, result.ColumnVolumeM3, result.ColumnHeightM);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in column calculation", ServiceName);
                result.Errors.Add($"Column calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Foundations

        private void CalculateFoundations(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                var foundations = new FilteredElementCollector(doc, selectedIds)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType()
                    .ToList();

                double totalVolume = 0;
                double totalArea = 0;

                foreach (var foundation in foundations)
                {
                    try
                    {
                        var volumeParam = foundation.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                        if (volumeParam != null && volumeParam.HasValue && volumeParam.AsDouble() > 0)
                            totalVolume += volumeParam.AsDouble();

                        var areaParam = foundation.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                        if (areaParam != null && areaParam.HasValue && areaParam.AsDouble() > 0)
                            totalArea += areaParam.AsDouble();
                        else
                        {
                            // Fallback: bounding box footprint
                            var bbox = foundation.get_BoundingBox(null);
                            if (bbox != null)
                                totalArea += Math.Abs(bbox.Max.X - bbox.Min.X) * Math.Abs(bbox.Max.Y - bbox.Min.Y);
                        }

                        result.FoundationCount++;
                    }
                    catch (Exception ex)
                    {
                        Log.Warning("[{Service}] Error calculating foundation {Id}: {Error}",
                            ServiceName, foundation.Id.Value, ex.Message);
                    }
                }

                result.FoundationVolumeM3 = ConvertToCubicMeters(totalVolume);
                result.FoundationAreaM2 = ConvertToSquareMeters(totalArea);

                Log.Information("[{Service}] Foundations: {Count}, Volume: {Vol:F2} m3, Area: {Area:F2} m2",
                    ServiceName, result.FoundationCount, result.FoundationVolumeM3, result.FoundationAreaM2);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in foundation calculation", ServiceName);
                result.Errors.Add($"Foundation calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Stairs

        private void CalculateStairs(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                var options = new Options
                {
                    ComputeReferences = true,
                    DetailLevel = ViewDetailLevel.Fine,
                    IncludeNonVisibleObjects = true
                };

                double totalStairVolume = 0;
                double totalRunTreadLength = 0;
                double totalLandingArea = 0;
                double totalTreadLength = 0;

                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category == null) continue;

                    long catId = element.Category.Id.Value;

                    // Parent Stairs: Volume (concrete)
                    if (catId == (long)BuiltInCategory.OST_Stairs)
                    {
                        try
                        {
                            result.StairCount++;

                            var riserParam = element.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_NUM_RISERS);
                            if (riserParam != null && riserParam.HasValue)
                                result.StairRiserCount += riserParam.AsInteger();

                            var geomElement = element.get_Geometry(options);
                            double volume = GetVolumeFromGeometry(geomElement);
                            totalStairVolume += volume;

                            // Hypothetical tread length
                            if (element is Stairs stairElement)
                            {
                                int riserCount = riserParam?.AsInteger() ?? 0;
                                int treadCount = Math.Max(0, riserCount - 1);

                                double treadDepth = 0;
                                var runIds = stairElement.GetStairsRuns();
                                if (runIds != null && runIds.Count > 0)
                                {
                                    var firstRun = doc.GetElement(runIds.First());
                                    if (firstRun != null)
                                    {
                                        var treadDepthParam = firstRun.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                                        if (treadDepthParam != null && treadDepthParam.HasValue)
                                            treadDepth = treadDepthParam.AsDouble();
                                    }
                                }

                                if (treadDepth <= 0)
                                {
                                    var parentTreadParam = element.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH)
                                                        ?? element.LookupParameter("Actual Tread Depth");
                                    if (parentTreadParam != null && parentTreadParam.HasValue && parentTreadParam.StorageType == StorageType.Double)
                                        treadDepth = parentTreadParam.AsDouble();
                                }

                                if (treadCount > 0 && treadDepth > 0)
                                    totalTreadLength += treadCount * treadDepth;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Warning("[{Service}] Error calculating stair {Id}: {Error}",
                                ServiceName, id.Value, ex.Message);
                        }
                    }
                    // Stair Runs: Tread length
                    else if (catId == (long)BuiltInCategory.OST_StairsRuns)
                    {
                        try
                        {
                            result.StairRunCount++;
                            double treadLength = 0;

                            var actualTreadDepth = element.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_TREAD_DEPTH);
                            var numRisers = element.get_Parameter(BuiltInParameter.STAIRS_ACTUAL_NUM_RISERS);

                            if (actualTreadDepth != null && actualTreadDepth.HasValue &&
                                numRisers != null && numRisers.HasValue)
                            {
                                treadLength = actualTreadDepth.AsDouble() * numRisers.AsInteger();
                            }

                            if (treadLength <= 0)
                            {
                                var lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                                if (lengthParam != null && lengthParam.HasValue && lengthParam.AsDouble() > 0)
                                    treadLength = lengthParam.AsDouble();
                            }

                            totalRunTreadLength += treadLength;
                        }
                        catch (Exception ex)
                        {
                            Log.Warning("[{Service}] Error calculating stair run {Id}: {Error}",
                                ServiceName, id.Value, ex.Message);
                        }
                    }
                    // Stair Landings: Area
                    else if (catId == (long)BuiltInCategory.OST_StairsLandings)
                    {
                        try
                        {
                            result.LandingCount++;
                            double area = 0;

                            var areaParam = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                            if (areaParam != null && areaParam.HasValue && areaParam.AsDouble() > 0)
                                area = areaParam.AsDouble();

                            if (area <= 0)
                            {
                                var bbox = element.get_BoundingBox(null);
                                if (bbox != null)
                                    area = Math.Abs(bbox.Max.X - bbox.Min.X) * Math.Abs(bbox.Max.Y - bbox.Min.Y);
                            }

                            totalLandingArea += area;
                        }
                        catch (Exception ex)
                        {
                            Log.Warning("[{Service}] Error calculating landing {Id}: {Error}",
                                ServiceName, id.Value, ex.Message);
                        }
                    }
                }

                result.StairVolumeM3 = ConvertToCubicMeters(totalStairVolume);
                result.StairTreadLengthM = ConvertToMeters(totalTreadLength);
                result.StairRunTreadLengthM = ConvertToMeters(totalRunTreadLength);
                result.LandingAreaM2 = ConvertToSquareMeters(totalLandingArea);

                Log.Information("[{Service}] Stairs: {Count} ({Risers} risers), Vol: {Vol:F2} m3 | Runs: {RC}, TreadLen: {TL:F2} m | Landings: {LC}, Area: {LA:F2} m2",
                    ServiceName, result.StairCount, result.StairRiserCount, result.StairVolumeM3,
                    result.StairRunCount, result.StairRunTreadLengthM,
                    result.LandingCount, result.LandingAreaM2);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in stair calculation", ServiceName);
                result.Errors.Add($"Stair calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Railings

        private void CalculateRailings(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                int railingCount = 0;
                double totalLength = 0;

                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category == null) continue;

                    long catId = element.Category.Id.Value;
                    if (catId == (long)BuiltInCategory.OST_Railings ||
                        catId == (long)BuiltInCategory.OST_StairsRailing)
                    {
                        railingCount++;
                        double length = 0;

                        // Method 1: Railing.GetPath()
                        if (element is Railing railing)
                        {
                            try
                            {
                                var path = railing.GetPath();
                                if (path != null && path.Count > 0)
                                {
                                    foreach (var curve in path)
                                    {
                                        if (curve != null)
                                            length += curve.Length;
                                    }
                                }
                            }
                            catch { /* fallback below */ }
                        }

                        // Method 2: Parameters
                        if (length <= 0)
                        {
                            var paramBips = new[]
                            {
                                BuiltInParameter.CURVE_ELEM_LENGTH,
                                BuiltInParameter.HOST_PERIMETER_COMPUTED
                            };

                            foreach (var bip in paramBips)
                            {
                                var param = element.get_Parameter(bip);
                                if (param != null && param.HasValue && param.AsDouble() > 0)
                                {
                                    length = param.AsDouble();
                                    break;
                                }
                            }
                        }

                        // Method 3: LocationCurve
                        if (length <= 0 && element.Location is LocationCurve locCurve && locCurve.Curve != null)
                            length = locCurve.Curve.Length;

                        // Method 4: BoundingBox diagonal
                        if (length <= 0)
                        {
                            var bbox = element.get_BoundingBox(null);
                            if (bbox != null)
                            {
                                double dx = Math.Abs(bbox.Max.X - bbox.Min.X);
                                double dy = Math.Abs(bbox.Max.Y - bbox.Min.Y);
                                length = Math.Sqrt(dx * dx + dy * dy);
                            }
                        }

                        totalLength += length;
                    }
                }

                result.RailingCount = railingCount;
                result.RailingLengthM = ConvertToMeters(totalLength);

                Log.Information("[{Service}] Railings: {Count}, Length: {Len:F2} m",
                    ServiceName, result.RailingCount, result.RailingLengthM);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in railing calculation", ServiceName);
                result.Errors.Add($"Railing calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Doors

        private void CalculateDoors(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category != null &&
                        element.Category.Id.Value == (long)BuiltInCategory.OST_Doors)
                    {
                        result.DoorCount++;
                    }
                }

                Log.Information("[{Service}] Doors: {Count}", ServiceName, result.DoorCount);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in door calculation", ServiceName);
                result.Errors.Add($"Door calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Windows

        private void CalculateWindows(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    if (element?.Category != null &&
                        element.Category.Id.Value == (long)BuiltInCategory.OST_Windows)
                    {
                        result.WindowCount++;
                    }
                }

                Log.Information("[{Service}] Windows: {Count}", ServiceName, result.WindowCount);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[{Service}] Error in window calculation", ServiceName);
                result.Errors.Add($"Window calculation error: {ex.Message}");
            }
        }

        #endregion

        #region Material Volumes

        private void CalculateMaterialVolumes(Document doc, ICollection<ElementId> selectedIds, ElementCalculationResult result)
        {
            try
            {
                var materialVolumes = new Dictionary<string, double>();

                foreach (var elementId in selectedIds)
                {
                    var element = doc.GetElement(elementId);
                    if (element == null) continue;

                    try
                    {
                        var materialIds = element.GetMaterialIds(false);
                        foreach (var materialId in materialIds)
                        {
                            var material = doc.GetElement(materialId) as Material;
                            if (material == null) continue;

                            string materialName = material.Name;
                            double volume = element.GetMaterialVolume(materialId);

                            if (volume > 0)
                            {
                                double volumeM3 = ConvertToCubicMeters(volume);
                                if (!materialVolumes.ContainsKey(materialName))
                                    materialVolumes[materialName] = 0;
                                materialVolumes[materialName] += volumeM3;
                            }
                        }
                    }
                    catch { /* skip element */ }
                }

                foreach (var kvp in materialVolumes.OrderByDescending(x => x.Value))
                {
                    result.VolumeByMaterial[kvp.Key] = kvp.Value;
                }

                Log.Debug("[{Service}] Found {Count} materials with volumes", ServiceName, materialVolumes.Count);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[{Service}] Error in material calculation", ServiceName);
            }
        }

        #endregion

        #region Geometry Helpers

        private double GetVolumeFromGeometry(GeometryElement geomElement, Transform transform = null)
        {
            double volume = 0.0;
            if (geomElement == null) return volume;

            foreach (var geomObj in geomElement)
            {
                if (geomObj is Solid solid && solid.Volume > 0)
                {
                    volume += solid.Volume;
                }
                else if (geomObj is GeometryInstance geomInstance)
                {
                    var instTransform = geomInstance.Transform;
                    if (transform != null)
                        instTransform = transform.Multiply(instTransform);

                    var instGeomElem = geomInstance.GetInstanceGeometry();
                    volume += GetVolumeFromGeometry(instGeomElem, instTransform);
                }
            }

            return volume;
        }

        #endregion

        #region Unit Conversion

        private static double ConvertToMeters(double internalValue)
        {
            return Math.Round(UnitUtils.ConvertFromInternalUnits(internalValue, UnitTypeId.Meters), 2);
        }

        private static double ConvertToSquareMeters(double internalValue)
        {
            return Math.Round(UnitUtils.ConvertFromInternalUnits(internalValue, UnitTypeId.SquareMeters), 2);
        }

        private static double ConvertToCubicMeters(double internalValue)
        {
            return Math.Round(UnitUtils.ConvertFromInternalUnits(internalValue, UnitTypeId.CubicMeters), 2);
        }

        #endregion
    }
}
