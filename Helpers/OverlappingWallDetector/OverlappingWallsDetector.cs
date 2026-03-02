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
        /// <summary>Overlap as percentage of the shorter wall (0–100).</summary>
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
        /// Perpendicular distance tolerance in feet (≈5 mm).
        /// Same-type walls with identical centerlines will have distance ≈ 0.
        /// </summary>
        private const double DistanceTolerance = 0.016;  // ~5 mm

        /// <summary>
        /// Minimum overlap length in feet to consider walls overlapping (≈50 mm).
        /// Prevents false positives from walls that only touch at endpoints.
        /// </summary>
        private const double MinOverlapLength = 0.164;  // ~50 mm

        /// <summary>
        /// Angular tolerance for parallel check (radians). ~1 degree.
        /// </summary>
        private const double AngleTolerance = 0.018;

        /// <summary>
        /// Finds groups of same-type walls that overlap each other.
        /// </summary>
        /// <param name="walls">Walls to analyze.</param>
        /// <returns>List of overlapping wall groups (each with 2+ walls).</returns>
        public static List<OverlappingWallGroup> FindOverlappingWalls(IList<Wall> walls)
        {
            // Group by WallType to only compare same-type walls
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
                    double maxPercent = 0;

                    for (int j = i + 1; j < wallList.Count; j++)
                    {
                        if (visited.Contains(wallList[j].Id))
                            continue;

                        double overlap = GetOverlapLength(wallList[i], wallList[j]);
                        if (overlap >= MinOverlapLength)
                        {
                            group.Add(wallList[j]);
                            visited.Add(wallList[j].Id);
                            maxOverlap = Math.Max(maxOverlap, overlap);

                            double shorterLen = Math.Min(
                                ((LocationCurve)wallList[i].Location).Curve.Length,
                                ((LocationCurve)wallList[j].Location).Curve.Length);
                            double pct = shorterLen > 0 ? (overlap / shorterLen) * 100.0 : 0;
                            maxPercent = Math.Max(maxPercent, pct);
                        }
                    }

                    if (group.Count >= 2)
                    {
                        visited.Add(wallList[i].Id);
                        result.Add(new OverlappingWallGroup
                        {
                            WallType = wallList[i].WallType,
                            Walls = group,
                            OverlapLength = maxOverlap,
                            OverlapPercent = maxPercent
                        });
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Returns the overlap length between two walls if they are nearly coincident,
        /// or 0 if they are not overlapping.
        /// </summary>
        private static double GetOverlapLength(Wall w1, Wall w2)
        {
            var curve1 = ((LocationCurve)w1.Location).Curve;
            var curve2 = ((LocationCurve)w2.Location).Curve;

            // Only handle linear walls (v1)
            if (curve1 is not Line line1 || curve2 is not Line line2)
                return 0;

            // Check parallel
            var dir1 = line1.Direction;
            var dir2 = line2.Direction;
            double cross = Math.Abs(dir1.X * dir2.Y - dir1.Y * dir2.X);
            if (cross > AngleTolerance)
                return 0;

            // Check perpendicular distance between the two lines
            var delta = line2.GetEndPoint(0) - line1.GetEndPoint(0);
            // Perpendicular component (in XY plane)
            var perp = delta - dir1.DotProduct(delta) * dir1;
            double perpDist = Math.Sqrt(perp.X * perp.X + perp.Y * perp.Y);
            if (perpDist > DistanceTolerance)
                return 0;

            // Check different base levels — same-type walls on different levels aren't overlapping
            var level1 = w1.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
            var level2 = w2.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId();
            if (level1 != null && level2 != null && level1 != level2)
                return 0;

            // Project both walls onto the same line direction to find overlap span
            double a1 = dir1.DotProduct(line1.GetEndPoint(0).IsAlmostEqualTo(XYZ.Zero)
                ? line1.GetEndPoint(0) : (XYZ)line1.GetEndPoint(0));
            double b1 = dir1.DotProduct((XYZ)line1.GetEndPoint(1));

            // For wall 2, project onto wall 1's direction (they're parallel)
            double a2 = dir1.DotProduct((XYZ)line2.GetEndPoint(0));
            double b2 = dir1.DotProduct((XYZ)line2.GetEndPoint(1));

            double min1 = Math.Min(a1, b1), max1 = Math.Max(a1, b1);
            double min2 = Math.Min(a2, b2), max2 = Math.Max(a2, b2);

            double overlapStart = Math.Max(min1, min2);
            double overlapEnd = Math.Min(max1, max2);
            double overlapLength = overlapEnd - overlapStart;

            return overlapLength > 0 ? overlapLength : 0;
        }
    }
}
