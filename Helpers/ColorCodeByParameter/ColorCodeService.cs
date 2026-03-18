using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BimTasksV2.Helpers.ColorCodeByParameter
{
    public class ColorGroupResult
    {
        public string PrimaryValue { get; set; } = "";
        public string SecondaryValue { get; set; } = "";
        public Color Color { get; set; } = new Color(128, 128, 128);
        public List<ElementId> ElementIds { get; set; } = new();
    }

    public static class ColorCodeService
    {
        private static readonly Color[] Palette = new[]
        {
            new Color(41, 128, 185),   // Blue
            new Color(231, 76, 60),    // Red
            new Color(39, 174, 96),    // Green
            new Color(243, 156, 18),   // Orange
            new Color(142, 68, 173),   // Purple
            new Color(22, 160, 133),   // Teal
            new Color(232, 67, 147),   // Pink
            new Color(160, 109, 55),   // Brown
            new Color(52, 73, 94),     // Dark Blue-Grey
            new Color(211, 84, 0),     // Dark Orange
            new Color(46, 204, 113),   // Emerald
            new Color(52, 152, 219),   // Light Blue
            new Color(155, 89, 182),   // Amethyst
            new Color(230, 126, 34),   // Carrot
            new Color(26, 188, 156),   // Turquoise
            new Color(192, 57, 43),    // Pomegranate
            new Color(44, 62, 80),     // Midnight Blue
            new Color(241, 196, 15),   // Sun Yellow
            new Color(127, 140, 141),  // Concrete Grey
            new Color(189, 195, 199),  // Silver
            new Color(99, 110, 114),   // Asbestos
            new Color(214, 48, 49),    // Alizarin
            new Color(108, 92, 231),   // Warm Purple
            new Color(0, 148, 50),     // Forest Green
        };

        public static List<string> GetParameterNames(Document doc, ICollection<ElementId> elementIds)
        {
            var names = new HashSet<string>();

            foreach (var id in elementIds)
            {
                var elem = doc.GetElement(id);
                if (elem == null) continue;

                foreach (Parameter p in elem.GetOrderedParameters())
                {
                    if (p?.Definition != null && p.HasValue &&
                        !string.IsNullOrWhiteSpace(p.Definition.Name))
                    {
                        names.Add(p.Definition.Name);
                    }
                }
            }

            return names.OrderBy(n => n).ToList();
        }

        public static Color[] GetPalette() => (Color[])Palette.Clone();

        private static double ColorDistanceSq(Color a, Color b)
        {
            double dr = a.Red - b.Red;
            double dg = a.Green - b.Green;
            double db = a.Blue - b.Blue;
            return dr * dr + dg * dg + db * db;
        }

        /// <summary>
        /// Returns <paramref name="count"/> colors with maximum distance between
        /// consecutive entries (greedy: each pick is most distant from the previous).
        /// </summary>
        public static List<Color> GetDistancedColors(int count)
        {
            if (count <= 0) return new();
            var available = new List<Color>(Palette);
            var result = new List<Color> { available[0] };
            available.RemoveAt(0);

            while (result.Count < count)
            {
                if (available.Count == 0)
                    available = new List<Color>(Palette);

                var last = result[^1];
                int bestIdx = 0;
                double bestDist = -1;
                for (int i = 0; i < available.Count; i++)
                {
                    double d = ColorDistanceSq(last, available[i]);
                    if (d > bestDist) { bestDist = d; bestIdx = i; }
                }
                result.Add(available[bestIdx]);
                available.RemoveAt(bestIdx);
            }
            return result;
        }

        /// <summary>
        /// Same as <see cref="GetDistancedColors"/> but with a random start
        /// and slight randomness (picks from top-3 most distant each step).
        /// </summary>
        public static List<Color> GetShuffledDistancedColors(int count)
        {
            if (count <= 0) return new();
            var rng = new Random();
            var available = new List<Color>(Palette);
            var result = new List<Color>();

            int startIdx = rng.Next(available.Count);
            result.Add(available[startIdx]);
            available.RemoveAt(startIdx);

            while (result.Count < count)
            {
                if (available.Count == 0)
                    available = new List<Color>(Palette);

                var last = result[^1];
                var ranked = new List<(int Index, double Dist)>();
                for (int i = 0; i < available.Count; i++)
                    ranked.Add((i, ColorDistanceSq(last, available[i])));
                ranked.Sort((a, b) => b.Dist.CompareTo(a.Dist));

                int topN = Math.Min(3, ranked.Count);
                var pick = ranked[rng.Next(topN)];
                result.Add(available[pick.Index]);
                available.RemoveAt(pick.Index);
            }
            return result;
        }

        private static FillPatternElement? GetSolidFillPattern(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(fp => fp.GetFillPattern().IsSolidFill);
        }

        public static void ApplyGroupOverride(
            Document doc, View view, ICollection<ElementId> elementIds, Color color)
        {
            var solidFill = GetSolidFillPattern(doc);
            if (solidFill == null) return;

            using var tx = new Transaction(doc, "Update Color Override");
            tx.Start();
            var ogs = new OverrideGraphicSettings();
            ogs.SetSurfaceForegroundPatternId(solidFill.Id);
            ogs.SetSurfaceForegroundPatternColor(color);
            foreach (var id in elementIds)
                view.SetElementOverrides(id, ogs);
            tx.Commit();
        }

        public static void ApplyMultipleGroupOverrides(
            Document doc, View view,
            List<(ICollection<ElementId> ElementIds, Color Color)> groups)
        {
            var solidFill = GetSolidFillPattern(doc);
            if (solidFill == null) return;

            using var tx = new Transaction(doc, "Shuffle Colors");
            tx.Start();
            foreach (var (elementIds, color) in groups)
            {
                var ogs = new OverrideGraphicSettings();
                ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                ogs.SetSurfaceForegroundPatternColor(color);
                foreach (var id in elementIds)
                    view.SetElementOverrides(id, ogs);
            }
            tx.Commit();
        }

        public static List<ColorGroupResult> ApplyColorOverrides(
            Document doc, View view, string primaryParam, string? secondaryParam,
            ICollection<ElementId> elementIds)
        {
            var solidFill = GetSolidFillPattern(doc);

            if (solidFill == null) return new List<ColorGroupResult>();

            // Group elements by primary parameter value
            var groups = new Dictionary<string, List<ElementId>>();
            foreach (var id in elementIds)
            {
                var elem = doc.GetElement(id);
                if (elem == null) continue;

                var param = elem.LookupParameter(primaryParam);
                var value = param != null && param.HasValue
                    ? param.AsValueString() ?? param.AsString() ?? "(empty)"
                    : "(no value)";

                if (!groups.ContainsKey(value))
                    groups[value] = new List<ElementId>();
                groups[value].Add(id);
            }

            var results = new List<ColorGroupResult>();
            var orderedGroups = groups.OrderBy(g => g.Key).ToList();
            var colorSequence = GetDistancedColors(orderedGroups.Count);
            int colorIndex = 0;

            using (var tx = new Transaction(doc, "Color Code by Parameter"))
            {
                tx.Start();

                foreach (var kvp in orderedGroups)
                {
                    var color = colorSequence[colorIndex++];

                    // Get secondary value from first element in group
                    string secondaryValue = "";
                    if (!string.IsNullOrEmpty(secondaryParam))
                    {
                        var firstElem = doc.GetElement(kvp.Value[0]);
                        if (firstElem != null)
                        {
                            var sp = firstElem.LookupParameter(secondaryParam);
                            if (sp != null && sp.HasValue)
                                secondaryValue = sp.AsValueString() ?? sp.AsString() ?? "";
                        }
                    }

                    // Apply override to each element
                    var ogs = new OverrideGraphicSettings();
                    ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                    ogs.SetSurfaceForegroundPatternColor(color);

                    foreach (var elemId in kvp.Value)
                    {
                        view.SetElementOverrides(elemId, ogs);
                    }

                    results.Add(new ColorGroupResult
                    {
                        PrimaryValue = kvp.Key,
                        SecondaryValue = secondaryValue,
                        Color = color,
                        ElementIds = kvp.Value
                    });
                }

                tx.Commit();
            }

            return results;
        }

        public static void ClearAllOverrides(Document doc, View view, ICollection<ElementId> elementIds)
        {
            using var tx = new Transaction(doc, "Reset Color Overrides");
            tx.Start();

            var empty = new OverrideGraphicSettings();
            foreach (var id in elementIds)
            {
                view.SetElementOverrides(id, empty);
            }

            tx.Commit();
        }

        public static void SelectElements(UIDocument uidoc, ICollection<ElementId> ids)
        {
            uidoc.Selection.SetElementIds(ids);
        }
    }
}
