using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using BimTasksV2.Commands.Handlers;
using Serilog;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Generic legend row: color swatch + primary label + optional secondary label.
    /// Used by both Filter-to-Legend and Color Swatch commands.
    /// </summary>
    public class LegendEntry
    {
        public Color Color { get; set; }
        public string Label { get; set; } = "";
        public string? SecondaryLabel { get; set; }
    }

    public class LegendBuilder
    {
        // Standard legend dimensions in mm
        private const double SwatchWidthMm = 12.0;
        private const double SwatchHeightMm = 6.0;
        private const double TextHeightMm = 2.5;
        private const double GapSwatchToTextMm = 4.0;
        private const double RowPitchMm = 12.0; // vertical distance between row starts
        private const double SecondaryTextOffsetMm = 50.0; // X offset from origin for secondary label

        public void UpdateLegend(Document doc, View legendView, List<FilterLegendItem> filters)
        {
            using (var t = new Transaction(doc, "Clear Legend Content"))
            {
                t.Start();

                var ownedElements = new FilteredElementCollector(doc, legendView.Id)
                    .WhereElementIsNotElementType()
                    .ToElementIds()
                    .ToList();

                foreach (var id in ownedElements)
                {
                    try { doc.Delete(id); }
                    catch { }
                }

                t.Commit();
            }

            using (var t = new Transaction(doc, "Populate Legend"))
            {
                t.Start();
                PlaceLegendContent(doc, legendView, filters);
                t.Commit();
            }
        }

        public View? CreateLegend(Document doc, string sourceViewName, List<FilterLegendItem> filters)
        {
            View? legendView = null;

            using (var t = new Transaction(doc, "Create Legend View"))
            {
                t.Start();
                legendView = CreateLegendView(doc, sourceViewName);
                t.Commit();
            }

            if (legendView == null)
            {
                Log.Error("[LegendBuilder] Failed to create legend view");
                return null;
            }

            using (var t = new Transaction(doc, "Populate Legend"))
            {
                t.Start();
                PlaceLegendContent(doc, legendView, filters);
                t.Commit();
            }

            return legendView;
        }

        /// <summary>
        /// Creates a legend/drafting view from generic LegendEntry items (two-column text).
        /// </summary>
        public View? CreateLegend(Document doc, string legendName, List<LegendEntry> entries)
        {
            View? legendView = null;

            using (var t = new Transaction(doc, "Create Legend View"))
            {
                t.Start();
                legendView = CreateLegendView(doc, legendName);
                t.Commit();
            }

            if (legendView == null)
            {
                Log.Error("[LegendBuilder] Failed to create legend view");
                return null;
            }

            using (var t = new Transaction(doc, "Populate Legend"))
            {
                t.Start();
                PlaceLegendContent(doc, legendView, entries);
                t.Commit();
            }

            return legendView;
        }

        private void PlaceLegendContent(Document doc, View legendView, List<LegendEntry> entries)
        {
            int scale = legendView.Scale;
            double swatchW = Mm(SwatchWidthMm) * scale;
            double swatchH = Mm(SwatchHeightMm) * scale;
            double textX = swatchW + Mm(GapSwatchToTextMm) * scale;
            double descTextX = Mm(SecondaryTextOffsetMm) * scale;
            double rowPitch = Mm(RowPitchMm) * scale;

            var textType = EnsureTextNoteType(doc);
            if (textType == null)
                Log.Warning("[LegendBuilder] Could not create TextNoteType — text will not be placed");

            bool hasSecondary = entries.Any(e => !string.IsNullOrEmpty(e.SecondaryLabel));

            Log.Information("[LegendBuilder] Placing {Count} legend entries in view '{ViewName}' (Scale=1:{Scale})",
                entries.Count, legendView.Name, scale);

            for (int i = 0; i < entries.Count; i++)
            {
                var entry = entries[i];
                double rowBottom = -(i * rowPitch);

                // --- Color swatch ---
                try
                {
                    var regionType = EnsureFilledRegionType(doc, entry.Color);
                    if (regionType != null)
                    {
                        var loop = new CurveLoop();
                        var p0 = new XYZ(0, rowBottom, 0);
                        var p1 = new XYZ(swatchW, rowBottom, 0);
                        var p2 = new XYZ(swatchW, rowBottom + swatchH, 0);
                        var p3 = new XYZ(0, rowBottom + swatchH, 0);

                        loop.Append(Line.CreateBound(p0, p1));
                        loop.Append(Line.CreateBound(p1, p2));
                        loop.Append(Line.CreateBound(p2, p3));
                        loop.Append(Line.CreateBound(p3, p0));

                        FilledRegion.Create(doc, regionType.Id, legendView.Id, new List<CurveLoop> { loop });
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[LegendBuilder] FilledRegion.Create FAILED for entry '{Label}'", entry.Label);
                }

                // --- Primary label ---
                if (textType != null)
                {
                    try
                    {
                        var textPos = new XYZ(textX, rowBottom + (swatchH / 2.0), 0);
                        TextNote.Create(doc, legendView.Id, textPos, entry.Label, textType.Id);
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "[LegendBuilder] TextNote.Create FAILED for '{Label}'", entry.Label);
                    }

                    // --- Secondary label ---
                    if (hasSecondary && !string.IsNullOrEmpty(entry.SecondaryLabel))
                    {
                        try
                        {
                            var descPos = new XYZ(descTextX, rowBottom + (swatchH / 2.0), 0);
                            TextNote.Create(doc, legendView.Id, descPos, entry.SecondaryLabel, textType.Id);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "[LegendBuilder] TextNote.Create FAILED for secondary '{Label}'", entry.SecondaryLabel);
                        }
                    }
                }
            }
        }

        private void PlaceLegendContent(Document doc, View legendView, List<FilterLegendItem> filters)
        {
            // Legend/drafting view coordinates must be multiplied by view scale
            // to produce the desired paper-space dimensions
            int scale = legendView.Scale;
            double swatchW = Mm(SwatchWidthMm) * scale;
            double swatchH = Mm(SwatchHeightMm) * scale;
            double textX = swatchW + Mm(GapSwatchToTextMm) * scale;
            double rowPitch = Mm(RowPitchMm) * scale;

            // Prepare types once before the loop
            var textType = EnsureTextNoteType(doc);
            if (textType == null)
                Log.Warning("[LegendBuilder] Could not create TextNoteType — text will not be placed");

            Log.Information("[LegendBuilder] Placing {Count} legend items in view '{ViewName}' (ViewType={ViewType}, Scale=1:{Scale})",
                filters.Count, legendView.Name, legendView.ViewType, scale);

            for (int i = 0; i < filters.Count; i++)
            {
                double rowBottom = -(i * rowPitch);

                // --- Color swatch ---
                try
                {
                    var regionType = EnsureFilledRegionType(doc, filters[i].OverrideColor);
                    if (regionType != null)
                    {
                        var loop = new CurveLoop();
                        var p0 = new XYZ(0, rowBottom, 0);
                        var p1 = new XYZ(swatchW, rowBottom, 0);
                        var p2 = new XYZ(swatchW, rowBottom + swatchH, 0);
                        var p3 = new XYZ(0, rowBottom + swatchH, 0);

                        loop.Append(Line.CreateBound(p0, p1));
                        loop.Append(Line.CreateBound(p1, p2));
                        loop.Append(Line.CreateBound(p2, p3));
                        loop.Append(Line.CreateBound(p3, p0));

                        FilledRegion.Create(doc, regionType.Id, legendView.Id, new List<CurveLoop> { loop });
                    }
                    else
                    {
                        Log.Warning("[LegendBuilder] No FilledRegionType for filter '{Name}' — skipping swatch",
                            filters[i].FilterName);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[LegendBuilder] FilledRegion.Create FAILED for filter '{Name}' in view type {ViewType}",
                        filters[i].FilterName, legendView.ViewType);
                }

                // --- Filter name text ---
                if (textType != null)
                {
                    try
                    {
                        // Vertically center text with swatch
                        var textPos = new XYZ(textX, rowBottom + (swatchH / 2.0), 0);
                        TextNote.Create(doc, legendView.Id, textPos, filters[i].FilterName, textType.Id);
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "[LegendBuilder] TextNote.Create FAILED for '{Name}'", filters[i].FilterName);
                    }
                }
            }
        }

        private View? CreateLegendView(Document doc, string sourceViewName)
        {
            View? legendView = null;

            // Try to duplicate a legend view named "empty"
            var emptyLegend = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => !v.IsTemplate &&
                    v.ViewType == ViewType.Legend &&
                    v.Name.Equals("empty", StringComparison.OrdinalIgnoreCase));

            if (emptyLegend != null)
            {
                try
                {
                    var duplicatedId = emptyLegend.Duplicate(ViewDuplicateOption.Duplicate);
                    legendView = doc.GetElement(duplicatedId) as View;
                    Log.Information("[LegendBuilder] Duplicated legend 'empty' → new legend view");
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[LegendBuilder] Failed to duplicate 'empty' legend — falling back to drafting view");
                }
            }
            else
            {
                Log.Information("[LegendBuilder] No legend named 'empty' found — will create drafting view");
            }

            // Fallback: drafting view
            if (legendView == null)
            {
                var draftingType = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Drafting);

                if (draftingType == null)
                {
                    Log.Error("[LegendBuilder] No Drafting ViewFamilyType found");
                    return null;
                }

                legendView = ViewDrafting.Create(doc, draftingType.Id);
            }

            // Unique name
            string baseName = $"Legend - {sourceViewName}";
            string name = baseName;
            int suffix = 2;

            while (ViewNameExists(doc, name))
            {
                name = $"{baseName} ({suffix})";
                suffix++;
            }

            legendView.Name = name;
            return legendView;
        }

        private static bool ViewNameExists(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Any(v => v.Name == name);
        }

        /// <summary>
        /// Get or create a FilledRegionType with a solid fill in the given color.
        /// </summary>
        private FilledRegionType? EnsureFilledRegionType(Document doc, Color color)
        {
            if (!color.IsValid)
            {
                Log.Warning("[LegendBuilder] Skipping swatch — color is invalid/uninitialized");
                return null;
            }

            string typeName = $"Legend #{color.Red:X2}{color.Green:X2}{color.Blue:X2}";

            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault(frt => frt.Name == typeName);

            if (existing != null) return existing;

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .FirstOrDefault();

            if (baseType == null)
            {
                Log.Error("[LegendBuilder] No FilledRegionType exists in document to duplicate from");
                return null;
            }

            var solidFillId = GetSolidFillPatternId(doc);
            if (solidFillId == ElementId.InvalidElementId)
            {
                Log.Error("[LegendBuilder] No solid fill pattern found in document");
                return null;
            }

            var newType = baseType.Duplicate(typeName) as FilledRegionType;
            if (newType == null) return null;

            // Solid foreground fill with the target color
            newType.ForegroundPatternId = solidFillId;
            newType.ForegroundPatternColor = color;

            // No background pattern — just the solid color
            newType.BackgroundPatternId = ElementId.InvalidElementId;

            // Not masking — we want the fill color to be visible
            newType.IsMasking = false;

            Log.Information("[LegendBuilder] Created FilledRegionType '{TypeName}' with color ({R},{G},{B})",
                typeName, color.Red, color.Green, color.Blue);

            return newType;
        }

        private static ElementId GetSolidFillPatternId(Document doc)
        {
            var solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);

            return solidFill?.Id ?? ElementId.InvalidElementId;
        }

        /// <summary>
        /// Get or create a TextNoteType at the standard legend text size.
        /// </summary>
        private TextNoteType? EnsureTextNoteType(Document doc)
        {
            string typeName = $"Legend Text {TextHeightMm:F1}mm";

            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault(t => t.Name == typeName);

            if (existing != null) return existing;

            var baseType = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .FirstOrDefault();

            if (baseType == null)
            {
                Log.Error("[LegendBuilder] No TextNoteType exists in document to duplicate from");
                return null;
            }

            var newType = baseType.Duplicate(typeName) as TextNoteType;
            if (newType == null) return null;

            double textSizeFeet = Mm(TextHeightMm);
            var sizeParam = newType.get_Parameter(BuiltInParameter.TEXT_SIZE);
            if (sizeParam != null && !sizeParam.IsReadOnly)
            {
                sizeParam.Set(textSizeFeet);
                Log.Information("[LegendBuilder] Created TextNoteType '{TypeName}' at {SizeMm}mm ({SizeFt:F5} ft)",
                    typeName, TextHeightMm, textSizeFeet);
            }
            else
            {
                Log.Warning("[LegendBuilder] TEXT_SIZE parameter is null or read-only on type '{TypeName}'", typeName);
            }

            return newType;
        }

        private static double Mm(double mm)
        {
            return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters);
        }
    }
}
