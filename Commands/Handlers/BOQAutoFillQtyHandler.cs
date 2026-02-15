using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Auto-fills BI_QtyBasis based on category and calculates BI_QtyValue from geometry.
    /// Processes all elements marked as pay items (BI_IsPayItem = Yes).
    /// Supports AREA, VOLUME, LENGTH, and COUNT basis types.
    /// Respects BI_QtyOverride and BI_QtyMultiplier if set.
    /// Also writes BI_HostElementId for hosted elements (doors, windows, panels, mullions).
    /// </summary>
    public class BOQAutoFillQtyHandler : ICommandHandler
    {
        private static readonly Dictionary<BuiltInCategory, string> CategoryDefaults =
            new Dictionary<BuiltInCategory, string>
            {
                { BuiltInCategory.OST_Floors, "AREA" },
                { BuiltInCategory.OST_Walls, "AREA" },
                { BuiltInCategory.OST_Roofs, "AREA" },
                { BuiltInCategory.OST_Ceilings, "AREA" },
                { BuiltInCategory.OST_StructuralFoundation, "VOLUME" },
                { BuiltInCategory.OST_StructuralColumns, "VOLUME" },
                { BuiltInCategory.OST_StructuralFraming, "VOLUME" },
                { BuiltInCategory.OST_StructuralFramingSystem, "VOLUME" },
                { BuiltInCategory.OST_Columns, "VOLUME" },
                { BuiltInCategory.OST_Doors, "COUNT" },
                { BuiltInCategory.OST_Windows, "COUNT" },
                { BuiltInCategory.OST_Stairs, "COUNT" },
                { BuiltInCategory.OST_StairsRailing, "LENGTH" },
                { BuiltInCategory.OST_Railings, "LENGTH" },
                { BuiltInCategory.OST_Ramps, "AREA" },
                { BuiltInCategory.OST_CurtainWallPanels, "AREA" },
                { BuiltInCategory.OST_CurtainWallMullions, "LENGTH" },
                { BuiltInCategory.OST_GenericModel, "COUNT" },
            };

        private static readonly BuiltInCategory[] SupportedCategories = CategoryDefaults.Keys.ToArray();

        // Revit internal units are feet
        private const double FeetToMeters = 0.3048;
        private const double SqFeetToSqMeters = FeetToMeters * FeetToMeters;
        private const double CuFeetToCuMeters = FeetToMeters * FeetToMeters * FeetToMeters;

        public void Execute(UIApplication uiApp)
        {
            Document doc = uiApp.ActiveUIDocument.Document;

            Log.Information("=== Starting BOQAutoFillQtyHandler ===");

            try
            {
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(CreateCategoryFilter());

                var allElements = collector.ToElements()
                    .Where(IsPayItem)
                    .ToList();

                if (allElements.Count == 0)
                {
                    TaskDialog.Show("BIMTasks - BOQ AutoFill",
                        "No pay items found.\n\nMake sure elements have BI_IsPayItem = Yes");
                    return;
                }

                Log.Information("Found {Count} pay items to process", allElements.Count);

                int updated = 0;
                int skipped = 0;
                var errors = new List<string>();

                using (Transaction tx = new Transaction(doc, "BOQ Auto-Fill Quantities"))
                {
                    tx.Start();

                    foreach (var elem in allElements)
                    {
                        try
                        {
                            bool wasUpdated = ProcessElement(elem);
                            if (wasUpdated) updated++;
                            else skipped++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"Element {elem.Id}: {ex.Message}");
                        }
                    }

                    tx.Commit();
                }

                ReportResults(allElements.Count, updated, skipped, errors);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "BOQAutoFillQtyHandler failed");
                TaskDialog.Show("BIMTasks - Error", $"An unexpected error occurred:\n\n{ex.Message}");
            }
        }

        private bool ProcessElement(Element elem)
        {
            string qtyBasis = GetParamString(elem, "BI_QtyBasis");

            if (string.IsNullOrWhiteSpace(qtyBasis))
            {
                qtyBasis = DetermineQtyBasis(elem);
                SetParamString(elem, "BI_QtyBasis", qtyBasis);
            }

            // Skip COMP items (PayLines) - they don't have physical quantities
            if (qtyBasis.ToUpper() == "COMP")
                return false;

            double qtyValue = CalculateQuantity(elem, qtyBasis);

            // Apply multiplier if set
            double multiplier = GetParamDouble(elem, "BI_QtyMultiplier");
            if (multiplier > 0 && multiplier != 1.0)
                qtyValue *= multiplier;

            // Check for override
            double qtyOverride = GetParamDouble(elem, "BI_QtyOverride");
            if (qtyOverride > 0)
                qtyValue = qtyOverride;

            SetParamDouble(elem, "BI_QtyValue", qtyValue);

            // Write host element reference
            string hostId = GetHostElementId(elem);
            if (hostId != null)
                SetParamString(elem, "BI_HostElementId", hostId);

            return true;
        }

        private string DetermineQtyBasis(Element elem)
        {
            var cat = (BuiltInCategory)elem.Category.Id.Value;

            // Special case: concrete floors/walls use VOLUME
            if (cat == BuiltInCategory.OST_Floors || cat == BuiltInCategory.OST_Walls)
            {
                string typeName = GetElementTypeName(elem);
                if (typeName != null && typeName.ToUpper().Contains("CONC"))
                    return "VOLUME";
            }

            if (CategoryDefaults.TryGetValue(cat, out string defaultBasis))
                return defaultBasis;

            return "COUNT";
        }

        private double CalculateQuantity(Element elem, string qtyBasis)
        {
            switch (qtyBasis.ToUpper())
            {
                case "AREA": return CalculateArea(elem);
                case "VOLUME": return CalculateVolume(elem);
                case "LENGTH": return CalculateLength(elem);
                case "COUNT":
                default: return 1.0;
            }
        }

        #region Quantity Calculations

        private double CalculateArea(Element elem)
        {
            double area = GetBuiltInParam(elem, BuiltInParameter.HOST_AREA_COMPUTED);
            if (area > 0) return ConvertToSquareMeters(area);
            return GetAreaFromGeometry(elem);
        }

        private double CalculateVolume(Element elem)
        {
            double volume = GetBuiltInParam(elem, BuiltInParameter.HOST_VOLUME_COMPUTED);
            if (volume > 0) return ConvertToCubicMeters(volume);
            return GetVolumeFromGeometry(elem);
        }

        private double CalculateLength(Element elem)
        {
            double length = GetBuiltInParam(elem, BuiltInParameter.CURVE_ELEM_LENGTH);
            if (length > 0) return ConvertToMeters(length);

            length = GetBuiltInParam(elem, BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH);
            if (length > 0) return ConvertToMeters(length);

            length = GetBuiltInParam(elem, BuiltInParameter.INSTANCE_LENGTH_PARAM);
            if (length > 0) return ConvertToMeters(length);

            if (elem.Location is LocationCurve locCurve)
                return ConvertToMeters(locCurve.Curve.Length);

            return 0;
        }

        private double GetAreaFromGeometry(Element elem)
        {
            try
            {
                var options = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Coarse };
                var geomElem = elem.get_Geometry(options);
                if (geomElem == null) return 0;

                double totalArea = 0;
                foreach (var geomObj in geomElem)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        foreach (Face face in solid.Faces)
                            totalArea += face.Area;
                    }
                }

                return ConvertToSquareMeters(totalArea / 6);
            }
            catch { return 0; }
        }

        private double GetVolumeFromGeometry(Element elem)
        {
            try
            {
                var options = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Coarse };
                var geomElem = elem.get_Geometry(options);
                if (geomElem == null) return 0;

                double totalVolume = 0;
                foreach (var geomObj in geomElem)
                {
                    if (geomObj is Solid solid)
                        totalVolume += solid.Volume;
                    else if (geomObj is GeometryInstance gi)
                    {
                        foreach (var instGeom in gi.GetInstanceGeometry())
                        {
                            if (instGeom is Solid instSolid)
                                totalVolume += instSolid.Volume;
                        }
                    }
                }

                return ConvertToCubicMeters(totalVolume);
            }
            catch { return 0; }
        }

        #endregion

        #region Unit Conversion

        private double ConvertToMeters(double feet) => Math.Round(feet * FeetToMeters, 4);
        private double ConvertToSquareMeters(double sqFeet) => Math.Round(sqFeet * SqFeetToSqMeters, 4);
        private double ConvertToCubicMeters(double cuFeet) => Math.Round(cuFeet * CuFeetToCuMeters, 4);

        #endregion

        #region Parameter Helpers

        private bool IsPayItem(Element elem)
        {
            var param = elem.LookupParameter("BI_IsPayItem");
            if (param == null || param.StorageType != StorageType.Integer) return false;
            return param.AsInteger() == 1;
        }

        private string GetParamString(Element elem, string paramName)
        {
            var param = elem.LookupParameter(paramName);
            if (param == null || param.StorageType != StorageType.String) return null;
            return param.AsString()?.Trim();
        }

        private double GetParamDouble(Element elem, string paramName)
        {
            var param = elem.LookupParameter(paramName);
            if (param == null || param.StorageType != StorageType.Double) return 0;
            return param.AsDouble();
        }

        private void SetParamString(Element elem, string paramName, string value)
        {
            var param = elem.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.String)
                param.Set(value ?? "");
        }

        private void SetParamDouble(Element elem, string paramName, double value)
        {
            var param = elem.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.Double)
                param.Set(value);
        }

        private string GetHostElementId(Element elem)
        {
            try
            {
                Element host = null;

                if (elem is Mullion mullion)
                    host = mullion.Host;
                else if (elem is Panel panel)
                    host = panel.Host;
                else if (elem is FamilyInstance fi)
                    host = fi.Host;

                return host?.UniqueId;
            }
            catch { return null; }
        }

        private double GetBuiltInParam(Element elem, BuiltInParameter bip)
        {
            var param = elem.get_Parameter(bip);
            if (param == null || param.StorageType != StorageType.Double) return 0;
            return param.AsDouble();
        }

        private string GetElementTypeName(Element elem)
        {
            var typeId = elem.GetTypeId();
            if (typeId == ElementId.InvalidElementId) return null;
            return elem.Document.GetElement(typeId)?.Name;
        }

        #endregion

        #region Filter & Reporting

        private ElementFilter CreateCategoryFilter()
        {
            var filters = SupportedCategories
                .Select(c => new ElementCategoryFilter(c))
                .Cast<ElementFilter>()
                .ToList();

            return new LogicalOrFilter(filters);
        }

        private void ReportResults(int total, int updated, int skipped, List<string> errors)
        {
            string msg = $"Processed {total} pay items:\n\n" +
                         $"Updated: {updated}\n" +
                         $"Skipped (COMP/unchanged): {skipped}\n";

            if (errors.Any())
            {
                msg += $"\nErrors: {errors.Count}\n";
                foreach (var e in errors.Take(5))
                    msg += $"  - {e}\n";
                if (errors.Count > 5)
                    msg += $"  ... and {errors.Count - 5} more\n";
            }

            TaskDialog.Show("BIMTasks - BOQ AutoFill", msg);
        }

        #endregion
    }
}
