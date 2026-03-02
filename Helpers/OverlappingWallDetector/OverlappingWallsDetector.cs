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
        public static List<OverlappingWallGroup> FindOverlappingWalls(IList<Wall> walls)
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

                        // Width overlap: how much of the wall width is shared
                        // wallWidth = total width. Two same-type walls overlap by (width - perpDist)
                        // When perpDist=0, overlap=100%. When perpDist=width, overlap=0%.
                        double wallWidth = wallList[i].Width;
                        double widthOverlap = wallWidth > 0
                            ? Math.Max(0, (wallWidth - minPerpDist) / wallWidth) * 100.0
                            : 100.0;

                        result.Add(new OverlappingWallGroup
                        {
                            WallType = wallList[i].WallType,
                            Walls = group,
                            OverlapLength = maxOverlap,
                            OverlapPercent = widthOverlap
                        });
                    }
                }
            }

            return result;
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
            double wallWidth = Math.Max(w1.Width, FallbackDistanceTolerance);
            if (perpDist >= wallWidth)
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
