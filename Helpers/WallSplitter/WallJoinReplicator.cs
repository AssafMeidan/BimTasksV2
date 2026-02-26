using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Records the join state of an original wall before splitting,
    /// then replicates that state onto the replacement walls.
    /// </summary>
    public class JoinRecord
    {
        public List<ElementId> GeometryJoinedElements { get; set; } = new();
        public bool AllowJoinAtStart { get; set; }
        public bool AllowJoinAtEnd { get; set; }
    }

    /// <summary>
    /// Records original wall join state and replicates it on replacement walls
    /// after a wall-split operation.
    /// </summary>
    public static class WallJoinReplicator
    {
        /// <summary>
        /// Captures the current join state of a wall before it is split/deleted.
        /// Records geometry-joined elements and end-join allow states.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="wall">The original wall about to be split.</param>
        /// <returns>A JoinRecord capturing the wall's join state.</returns>
        public static JoinRecord RecordJoins(Document doc, Wall wall)
        {
            var record = new JoinRecord();

            // Record geometry-joined elements
            var joinedIds = JoinGeometryUtils.GetJoinedElements(doc, wall);
            foreach (var id in joinedIds)
            {
                record.GeometryJoinedElements.Add(id);
            }

            // Record end-join allowed states
            record.AllowJoinAtStart = WallUtils.IsWallJoinAllowedAtEnd(wall, 0);
            record.AllowJoinAtEnd = WallUtils.IsWallJoinAllowedAtEnd(wall, 1);

            return record;
        }

        /// <summary>
        /// Replicates the original wall's join state onto the replacement walls.
        /// Joins adjacent replacement walls to each other, re-joins them to original
        /// neighbors, and toggles end joins on all replacements to force Revit corner cleanup.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="replacementWalls">Replacement walls in order (start-to-end of original).</param>
        /// <param name="record">The join record captured before splitting.</param>
        public static void ReplicateJoins(Document doc, List<Wall> replacementWalls, JoinRecord record)
        {
            if (replacementWalls == null || replacementWalls.Count == 0)
                return;

            // 1. Join adjacent replacement walls to each other (geometry join for overlapping layers)
            for (int i = 0; i < replacementWalls.Count - 1; i++)
            {
                try
                {
                    JoinGeometryUtils.JoinGeometry(doc, replacementWalls[i], replacementWalls[i + 1]);
                }
                catch
                {
                    Log.Warning(
                        "WallJoinReplicator: Failed to join adjacent replacement walls {WallA} and {WallB}",
                        replacementWalls[i].Id, replacementWalls[i + 1].Id);
                }
            }

            // 2. Re-join each replacement wall to original geometry-joined neighbors (if they still exist)
            foreach (var neighborId in record.GeometryJoinedElements)
            {
                var neighbor = doc.GetElement(neighborId);
                if (neighbor == null)
                    continue; // Neighbor was deleted (e.g., also split in batch mode)

                foreach (var replacement in replacementWalls)
                {
                    try
                    {
                        if (!JoinGeometryUtils.AreElementsJoined(doc, replacement, neighbor))
                        {
                            JoinGeometryUtils.JoinGeometry(doc, replacement, neighbor);
                        }
                    }
                    catch
                    {
                        Log.Warning(
                            "WallJoinReplicator: Failed to join replacement wall {WallId} to neighbor {NeighborId}",
                            replacement.Id, neighborId);
                    }
                }
            }

            // 3. Toggle end joins on ALL replacement walls at BOTH ends to force Revit
            //    to recalculate corner geometry. This makes Revit detect nearby walls
            //    and form clean L/T intersections.
            foreach (var wall in replacementWalls)
            {
                ForceEndJoinRecalculation(wall, 0);
                ForceEndJoinRecalculation(wall, 1);
            }

            // 4. Restore original end-join states on outermost ends
            //    (if original had joins disallowed, preserve that)
            if (!record.AllowJoinAtStart)
            {
                foreach (var wall in replacementWalls)
                    RestoreEndJoinState(wall, 0, false);
            }
            if (!record.AllowJoinAtEnd)
            {
                foreach (var wall in replacementWalls)
                    RestoreEndJoinState(wall, 1, false);
            }
        }

        /// <summary>
        /// Connects replacement walls from different original walls at corners
        /// using curve extension/trimming (same approach as CladdingWallJoinHelper).
        /// Matches walls by WallType first, then by layer position, so exterior
        /// connects to exterior and core connects to core.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="splitResults">Successful split results, each containing replacement pairs from one original wall.</param>
        public static void CrossJoinReplacements(Document doc, List<SplitResult> splitResults)
        {
            if (splitResults == null || splitResults.Count < 2)
                return;

            var connected = new HashSet<(long, long)>();

            // For each pair of original walls (different split results)
            for (int a = 0; a < splitResults.Count; a++)
            {
                for (int b = a + 1; b < splitResults.Count; b++)
                {
                    var pairsA = splitResults[a].ReplacementPairs;
                    var pairsB = splitResults[b].ReplacementPairs;

                    // Pass 1: Match by same WallType
                    foreach (var pa in pairsA)
                    {
                        foreach (var pb in pairsB)
                        {
                            var key = OrderIds(pa.Wall.Id, pb.Wall.Id);
                            if (connected.Contains(key)) continue;

                            if (pa.Wall.WallType.Id == pb.Wall.WallType.Id)
                            {
                                if (TrimWallsToCorner(doc, pa.Wall, pb.Wall))
                                    connected.Add(key);
                            }
                        }
                    }

                    // Pass 2: Match by same layer index (outer↔outer, inner↔inner)
                    foreach (var pa in pairsA)
                    {
                        foreach (var pb in pairsB)
                        {
                            var key = OrderIds(pa.Wall.Id, pb.Wall.Id);
                            if (connected.Contains(key)) continue;

                            if (pa.Layer.Index == pb.Layer.Index)
                            {
                                if (TrimWallsToCorner(doc, pa.Wall, pb.Wall))
                                    connected.Add(key);
                            }
                        }
                    }
                }
            }

            // Allow end joins on all replacement walls to form clean corners
            foreach (var result in splitResults)
            {
                foreach (var pair in result.ReplacementPairs)
                {
                    ForceEndJoinRecalculation(pair.Wall, 0);
                    ForceEndJoinRecalculation(pair.Wall, 1);
                }
            }
        }

        // Tolerance for extending curves to detect near-intersections (in feet)
        private const double ExtendTolerance = 1.0 / 30.48 * 10.0; // ~10 cm

        /// <summary>
        /// Extends both wall curves slightly, checks for intersection, and trims
        /// both to the intersection point so they meet at the corner.
        /// Returns true if walls were trimmed.
        /// </summary>
        private static bool TrimWallsToCorner(Document doc, Wall wall1, Wall wall2)
        {
            try
            {
                var loc1 = wall1.Location as LocationCurve;
                var loc2 = wall2.Location as LocationCurve;
                if (loc1 == null || loc2 == null) return false;

                Curve ext1 = ExtendLine(loc1.Curve, ExtendTolerance);
                Curve ext2 = ExtendLine(loc2.Curve, ExtendTolerance);

                IntersectionResultArray resultArray;
                var result = ext1.Intersect(ext2, out resultArray);

                if (result != SetComparisonResult.Overlap || resultArray == null || resultArray.Size == 0)
                    return false;

                XYZ intersectionPoint = resultArray.get_Item(0).XYZPoint;

                // Trim both curves to the intersection point
                Curve trimmed1 = TrimToIntersection(loc1.Curve, intersectionPoint);
                Curve trimmed2 = TrimToIntersection(loc2.Curve, intersectionPoint);

                if (trimmed1 != null && trimmed2 != null)
                {
                    loc1.Curve = trimmed1;
                    loc2.Curve = trimmed2;
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "WallJoinReplicator: Failed to trim walls {W1} and {W2} to corner",
                    wall1.Id, wall2.Id);
            }

            return false;
        }

        /// <summary>
        /// Extends a line curve by a specified length in both directions.
        /// Non-line curves are returned unchanged.
        /// </summary>
        private static Curve ExtendLine(Curve curve, double extension)
        {
            if (curve is Line line)
            {
                XYZ start = line.GetEndPoint(0);
                XYZ end = line.GetEndPoint(1);
                XYZ dir = (end - start).Normalize();
                return Line.CreateBound(start - dir * extension, end + dir * extension);
            }
            return curve;
        }

        /// <summary>
        /// Trims a line to an intersection point — keeps the farther endpoint
        /// and replaces the nearer one with the intersection point.
        /// </summary>
        private static Curve TrimToIntersection(Curve curve, XYZ point)
        {
            if (curve is Line)
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);

                // Replace the endpoint closest to the intersection
                if (start.DistanceTo(point) < end.DistanceTo(point))
                    return Line.CreateBound(point, end);
                else
                    return Line.CreateBound(start, point);
            }
            return null;
        }

        /// <summary>
        /// Returns a canonical ordered pair for use as a HashSet key.
        /// </summary>
        private static (long, long) OrderIds(ElementId a, ElementId b)
        {
            return a.Value < b.Value ? (a.Value, b.Value) : (b.Value, a.Value);
        }

        /// <summary>
        /// Toggles end join off then on to force Revit to recalculate the
        /// wall join geometry at a specific end. This triggers Revit's automatic
        /// detection of nearby walls and forms clean L/T corner intersections.
        /// </summary>
        private static void ForceEndJoinRecalculation(Wall wall, int end)
        {
            try
            {
                WallUtils.DisallowWallJoinAtEnd(wall, end);
                WallUtils.AllowWallJoinAtEnd(wall, end);
            }
            catch
            {
                // Non-critical — wall may not support end joins
            }
        }

        /// <summary>
        /// Sets the allow/disallow join state on a specific end of a wall.
        /// </summary>
        /// <param name="wall">The wall to update.</param>
        /// <param name="end">The end index (0 = start, 1 = end).</param>
        /// <param name="allow">True to allow joining, false to disallow.</param>
        private static void RestoreEndJoinState(Wall wall, int end, bool allow)
        {
            try
            {
                if (allow)
                    WallUtils.AllowWallJoinAtEnd(wall, end);
                else
                    WallUtils.DisallowWallJoinAtEnd(wall, end);
            }
            catch
            {
                Log.Warning(
                    "WallJoinReplicator: Failed to set end-join state on wall {WallId} end {End}",
                    wall.Id, end);
            }
        }
    }
}
