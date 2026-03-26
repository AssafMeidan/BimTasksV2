using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers.OverlappingWallDetector
{
    /// <summary>
    /// A group of walls that overlap each other at the same location.
    /// All walls in a group share the same WallType.
    /// </summary>
    public class OverlappingWallGroup
    {
        public WallType WallType { get; set; } = null!;
        public List<Wall> Walls { get; set; } = new();
        public double OverlapLength { get; set; }
        /// <summary>Width overlap percentage (0–100). 100% = centerlines identical, less = offset.</summary>
        public double OverlapPercent { get; set; }
    }

    /// <summary>
    /// Detects walls of the same type that occupy the same physical space.
    /// Two same-type walls are "overlapping" when their centerline curves are nearly
    /// coincident (parallel, close together, and span the same length).
    /// Uses a fast 2D pre-filter, then a precise 3D solid intersection for final overlap %.
    /// </summary>
    public static class OverlappingWallsDetector
    {
        /// <summary>
        /// Fallback max distance if wall width is unavailable.
        /// </summary>
        private const double FallbackDistanceTolerance = 0.5;  // ~150 mm

        /// <summary>
        /// Minimum overlap length in feet to consider walls overlapping (≈50 mm).
        /// Prevents false positives from walls that only touch at endpoints.
        /// </summary>
        private const double MinOverlapLength = 0.164;  // ~50 mm

        /// <summary>
        /// Angular tolerance for parallel check (radians). ~1 degree.
        /// </summary>
        private const double AngleTolerance = 0.018;

        private struct OverlapResult
        {
            public double Length;
            public double PerpDistance;
            public bool IsOverlapping;
        }

        /// <summary>
        /// Finds groups of same-type walls that overlap each other.
        /// </summary>
        public static List<OverlappingWallGroup> FindOverlappingWalls(IList<Wall> walls, Document? doc = null)
        {
            var byType = walls
                .Where(w => w.Location is LocationCurve)
                .GroupBy(w => w.WallType.Id.Value)
                .Where(g => g.Count() >= 2);

            var result = new List<OverlappingWallGroup>();

            foreach (var typeGroup in byType)
            {
                var wallList = typeGroup.ToList();
                var visited = new HashSet<ElementId>();

                for (int i = 0; i < wallList.Count; i++)
                {
                    if (visited.Contains(wallList[i].Id))
                        continue;

                    var group = new List<Wall> { wallList[i] };
                    double maxOverlap = 0;
                    double minPerpDist = double.MaxValue;

                    for (int j = i + 1; j < wallList.Count; j++)
                    {
                        if (visited.Contains(wallList[j].Id))
                            continue;

                        var overlap = CheckOverlap(wallList[i], wallList[j]);
                        if (overlap.IsOverlapping && overlap.Length >= MinOverlapLength)
                        {
                            group.Add(wallList[j]);
                            visited.Add(wallList[j].Id);
                            maxOverlap = Math.Max(maxOverlap, overlap.Length);
                            minPerpDist = Math.Min(minPerpDist, overlap.PerpDistance);
                        }
                    }

                    if (group.Count >= 2)
                    {
                        visited.Add(wallList[i].Id);

                        double overlapPercent;

                        if (doc != null)
                        {
                            // 3D solid intersection for accurate volume-based overlap
                            overlapPercent = ComputeVolumeOverlapPercent(group, doc);
                        }
                        else
                        {
                            // Fallback: 2D width-based estimate
                            double wallWidth = wallList[i].Width;
                            overlapPercent = wallWidth > 0
                                ? Math.Max(0, (wallWidth - minPerpDist) / wallWidth) * 100.0
                                : 100.0;
                        }

                        // Only report groups with meaningful overlap
                        if (overlapPercent < 1.0)
                            continue;

                        result.Add(new OverlappingWallGroup
                        {
                            WallType = wallList[i].WallType,
                            Walls = group,
                            OverlapLength = maxOverlap,
                            OverlapPercent = overlapPercent
                        });
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Computes volume overlap percentage for a group of walls using 3D solid boolean intersection.
        /// Returns the max pairwise overlap % (intersection volume / smaller wall volume).
        /// </summary>
        private static double ComputeVolumeOverlapPercent(List<Wall> group, Document doc)
        {
            double maxPercent = 0;
            var options = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Coarse };

            for (int i = 0; i < group.Count; i++)
            {
                var solid1 = GetLargestSolid(group[i], options);
                if (solid1 == null || solid1.Volume < 1e-9)
                    continue;

                for (int j = i + 1; j < group.Count; j++)
                {
                    var solid2 = GetLargestSolid(group[j], options);
                    if (solid2 == null || solid2.Volume < 1e-9)
                        continue;

                    try
                    {
                        var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                            solid1, solid2, BooleanOperationsType.Intersect);

                        if (intersection != null && intersection.Volume > 1e-9)
                        {
                            double smallerVolume = Math.Min(solid1.Volume, solid2.Volume);
                            double pct = (intersection.Volume / smallerVolume) * 100.0;
                            maxPercent = Math.Max(maxPercent, Math.Min(pct, 100.0));
                        }
                    }
                    catch
                    {
                        // Boolean op can fail on degenerate geometry — fall back to 2D estimate
                        double wallWidth = group[i].Width;
                        if (wallWidth > 0)
                        {
                            var curve1 = ((LocationCurve)group[i].Location).Curve;
                            var curve2 = ((LocationCurve)group[j].Location).Curve;
                            if (curve1 is Line line1 && curve2 is Line line2)
                            {
                                var delta = line2.GetEndPoint(0) - line1.GetEndPoint(0);
                                var perp = delta - line1.Direction.DotProduct(delta) * line1.Direction;
                                double perpDist = Math.Sqrt(perp.X * perp.X + perp.Y * perp.Y);
                                double pct = Math.Max(0, (wallWidth - perpDist) / wallWidth) * 100.0;
                                maxPercent = Math.Max(maxPercent, pct);
                            }
                        }
                    }
                }
            }

            return maxPercent;
        }

        /// <summary>
        /// Extracts the largest solid from an element's geometry.
        /// </summary>
        private static Solid? GetLargestSolid(Element element, Options options)
        {
            var geom = element.get_Geometry(options);
            if (geom == null) return null;

            Solid? largest = null;
            double maxVol = 0;

            foreach (var obj in geom)
            {
                switch (obj)
                {
                    case Solid s when s.Volume > maxVol:
                        largest = s;
                        maxVol = s.Volume;
                        break;
                    case GeometryInstance gi:
                        foreach (var inner in gi.GetInstanceGeometry())
                        {
                            if (inner is Solid s2 && s2.Volume > maxVol)
                            {
                                largest = s2;
                                maxVol = s2.Volume;
                            }
                        }
                        break;
                }
            }

            return largest;
        }

        /// <summary>
        /// Checks if two walls overlap and returns overlap length + perpendicular distance.
        /// </summary>
        private static OverlapResult CheckOverlap(Wall w1, Wall w2)
        {
            var curve1 = ((LocationCurve)w1.Location).Curve;
            var curve2 = ((LocationCurve)w2.Location).Curve;

            // Only handle linear walls (v1)
            if (curve1 is not Line line1 || curve2 is not Line line2)
                return default;

            // Check parallel
            var dir1 = line1.Direction;
            var dir2 = line2.Direction;
            double cross = Math.Abs(dir1.X * dir2.Y - dir1.Y * dir2.X);
            if (cross > AngleTolerance)
                return default;

            // Perpendicular distance between the two lines
            // Must be strictly less than wall width — at exactly wallWidth they just touch edges
            var delta = line2.GetEndPoint(0) - line1.GetEndPoint(0);
            var perp = delta - dir1.DotProduct(delta) * dir1;
            double perpDist = Math.Sqrt(perp.X * perp.X + perp.Y * perp.Y);
            // Only flag walls overlapping by more than 50% of width
            // (perpDist < halfWidth means >50% width overlap)
            double wallWidth = Math.Max(w1.Width, FallbackDistanceTolerance);
            if (perpDist >= wallWidth * 0.5)
                return default;

            // Check different base levels
            var level1 = w1.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
            var level2 = w2.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
            if (level1 != null && level2 != null && level1 != level2)
                return default;

            // Project both walls onto the same direction to find overlap span
            double a1 = dir1.DotProduct(line1.GetEndPoint(0));
            double b1 = dir1.DotProduct(line1.GetEndPoint(1));
            double a2 = dir1.DotProduct(line2.GetEndPoint(0));
            double b2 = dir1.DotProduct(line2.GetEndPoint(1));

            double min1 = Math.Min(a1, b1), max1 = Math.Max(a1, b1);
            double min2 = Math.Min(a2, b2), max2 = Math.Max(a2, b2);

            double overlapStart = Math.Max(min1, min2);
            double overlapEnd = Math.Min(max1, max2);
            double overlapLength = overlapEnd - overlapStart;

            if (overlapLength <= 0)
                return default;

            return new OverlapResult
            {
                Length = overlapLength,
                PerpDistance = perpDist,
                IsOverlapping = true
            };
        }
    }
}
