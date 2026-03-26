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
        /// neighbors, and restores end-join allow/disallow states.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="replacementWalls">Replacement walls in order (start-to-end of original).</param>
        /// <param name="record">The join record captured before splitting.</param>
        public static void ReplicateJoins(Document doc, List<Wall> replacementWalls, JoinRecord record)
        {
            if (replacementWalls == null || replacementWalls.Count == 0)
                return;

            // 1. Join adjacent replacement walls to each other
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

            // 2. Re-join each replacement wall to original geometry-joined neighbors
            foreach (var neighborId in record.GeometryJoinedElements)
            {
                var neighbor = doc.GetElement(neighborId);
                if (neighbor == null)
                    continue; // Original wall was deleted or neighbor no longer exists

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

            // 3. Restore end-join states on the first and last replacement walls
            RestoreEndJoinState(replacementWalls[0], 0, record.AllowJoinAtStart);

            var lastWall = replacementWalls[replacementWalls.Count - 1];
            RestoreEndJoinState(lastWall, 1, record.AllowJoinAtEnd);
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

        // ===== Cross-Wall Corner Connection (between different original walls) =====

        // Max distance between endpoints to consider walls as corner candidates (feet)
        private static readonly double MaxEndpointDistance =
            UnitUtils.ConvertToInternalUnits(500, UnitTypeId.Millimeters);

        /// <summary>
        /// Trims/extends replacement walls so the end result looks the same as before splitting.
        /// Two passes:
        ///   1. Between split groups: match by layer index (outer↔outer, inner↔inner)
        ///   2. Each replacement wall against ALL nearby unsplit neighbor walls
        /// </summary>
        public static void CrossJoinReplacements(Document doc, List<SplitResult> splitResults)
        {
            if (splitResults == null || splitResults.Count == 0)
                return;

            // Re-fetch all replacement walls from the document
            var groups = new List<List<(Wall Wall, LayerInfo Layer)>>();
            var allReplacementIds = new HashSet<long>();

            foreach (var sr in splitResults)
            {
                var freshPairs = new List<(Wall Wall, LayerInfo Layer)>();
                foreach (var rid in sr.ReplacementIds)
                {
                    var w = doc.GetElement(rid.WallId) as Wall;
                    if (w != null && w.IsValidObject)
                    {
                        freshPairs.Add((w, rid.Layer));
                        allReplacementIds.Add(rid.WallId.Value);
                    }
                }
                if (freshPairs.Count > 0)
                    groups.Add(freshPairs);
            }

            var connected = new HashSet<(long, long)>();
            int trimCount = 0;

            // Pass 1: Between split groups — match by layer index
            // Layer index 0 = outermost (exterior), higher = toward interior.
            // This correctly pairs outer↔outer, concrete↔concrete, inner↔inner.
            if (groups.Count >= 2)
            {
                for (int a = 0; a < groups.Count; a++)
                {
                    for (int b = a + 1; b < groups.Count; b++)
                    {
                        foreach (var pa in groups[a])
                        {
                            foreach (var pb in groups[b])
                            {
                                if (pa.Layer.Index != pb.Layer.Index) continue;

                                var key = OrderIds(pa.Wall.Id, pb.Wall.Id);
                                if (connected.Contains(key)) continue;

                                if (TrimWallsToCorner(pa.Wall, pb.Wall))
                                {
                                    connected.Add(key);
                                    trimCount++;
                                    Log.Debug("[CrossJoin] Pass1: Trimmed layer {Idx} walls {W1} ↔ {W2}",
                                        pa.Layer.Index, pa.Wall.Id, pb.Wall.Id);
                                }
                            }
                        }
                    }
                }
            }

            // Pass 2: Each replacement wall against nearby unsplit neighbor walls
            var allReplacementWalls = groups.SelectMany(g => g.Select(p => p.Wall)).ToList();

            foreach (var repWall in allReplacementWalls)
            {
                var neighbors = FindNearbyWalls(doc, repWall, MaxEndpointDistance);
                foreach (var neighbor in neighbors)
                {
                    // Skip other replacement walls (handled by pass 1)
                    if (allReplacementIds.Contains(neighbor.Id.Value))
                        continue;

                    var key = OrderIds(repWall.Id, neighbor.Id);
                    if (connected.Contains(key)) continue;

                    if (TrimWallsToCorner(repWall, neighbor))
                    {
                        connected.Add(key);
                        trimCount++;
                        Log.Debug("[CrossJoin] Pass2: Trimmed wall {W1} ↔ neighbor {W2}",
                            repWall.Id, neighbor.Id);
                    }
                }
            }

            Log.Information("[CrossJoin] Trimmed {Count} wall pairs across {Groups} groups",
                trimCount, groups.Count);

            // Force end join recalculation on all replacement walls
            foreach (var wall in allReplacementWalls)
            {
                ForceEndJoinRecalculation(wall, 0);
                ForceEndJoinRecalculation(wall, 1);
            }
        }

        /// <summary>
        /// Finds walls in the document whose endpoints are within maxDistance
        /// of the given wall's endpoints.
        /// </summary>
        private static List<Wall> FindNearbyWalls(Document doc, Wall wall, double maxDistance)
        {
            var loc = wall.Location as LocationCurve;
            if (loc == null) return new List<Wall>();

            XYZ start = loc.Curve.GetEndPoint(0);
            XYZ end = loc.Curve.GetEndPoint(1);

            var result = new List<Wall>();

            // Use a BoundingBox filter to narrow candidates, then check endpoints
            var bb = wall.get_BoundingBox(null);
            if (bb == null) return result;

            var expandedMin = new XYZ(bb.Min.X - maxDistance, bb.Min.Y - maxDistance, bb.Min.Z - maxDistance);
            var expandedMax = new XYZ(bb.Max.X + maxDistance, bb.Max.Y + maxDistance, bb.Max.Z + maxDistance);
            var outline = new Outline(expandedMin, expandedMax);
            var bbFilter = new BoundingBoxIntersectsFilter(outline);

            var candidates = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WherePasses(bbFilter)
                .Cast<Wall>()
                .Where(w => w.Id != wall.Id);

            foreach (var candidate in candidates)
            {
                var candidateLoc = candidate.Location as LocationCurve;
                if (candidateLoc == null) continue;

                XYZ cs = candidateLoc.Curve.GetEndPoint(0);
                XYZ ce = candidateLoc.Curve.GetEndPoint(1);

                if (start.DistanceTo(cs) < maxDistance || start.DistanceTo(ce) < maxDistance ||
                    end.DistanceTo(cs) < maxDistance || end.DistanceTo(ce) < maxDistance)
                {
                    result.Add(candidate);
                }
            }

            return result;
        }

        /// <summary>
        /// Computes the mathematical intersection of two wall lines (unbounded),
        /// checks the intersection is near an endpoint of each wall, and trims
        /// both to the intersection point so they meet at the corner.
        /// Uses analytic 2D line-line intersection — works regardless of wall
        /// thickness or offset distance (no extension tolerance needed).
        /// </summary>
        private static bool TrimWallsToCorner(Wall wall1, Wall wall2)
        {
            try
            {
                var loc1 = wall1.Location as LocationCurve;
                var loc2 = wall2.Location as LocationCurve;
                if (loc1 == null || loc2 == null) return false;
                if (!(loc1.Curve is Line) || !(loc2.Curve is Line)) return false;

                XYZ p1 = loc1.Curve.GetEndPoint(0);
                XYZ p2 = loc1.Curve.GetEndPoint(1);
                XYZ p3 = loc2.Curve.GetEndPoint(0);
                XYZ p4 = loc2.Curve.GetEndPoint(1);

                XYZ d1 = p2 - p1;
                XYZ d2 = p4 - p3;

                // 2D cross product (XY plane) — zero means parallel
                double cross = d1.X * d2.Y - d1.Y * d2.X;
                if (Math.Abs(cross) < 1e-10)
                    return false;

                // Solve for parameter t: intersection = p1 + t * d1
                XYZ diff = p3 - p1;
                double t = (diff.X * d2.Y - diff.Y * d2.X) / cross;

                // Use Z from the wall (both should be at the same level)
                XYZ intersection = new XYZ(
                    p1.X + t * d1.X,
                    p1.Y + t * d1.Y,
                    p1.Z);

                // Only trim if the intersection is near an endpoint of EACH wall
                // (i.e., this is actually a corner, not two distant walls)
                double distToWall1 = Math.Min(intersection.DistanceTo(p1), intersection.DistanceTo(p2));
                double distToWall2 = Math.Min(intersection.DistanceTo(p3), intersection.DistanceTo(p4));

                if (distToWall1 > MaxEndpointDistance || distToWall2 > MaxEndpointDistance)
                    return false;

                // Replace the endpoint closest to the intersection on each wall
                Line trimmed1 = TrimLineToPoint(p1, p2, intersection);
                Line trimmed2 = TrimLineToPoint(p3, p4, intersection);

                loc1.Curve = trimmed1;
                loc2.Curve = trimmed2;

                Log.Debug("[TrimCorner] {W1}↔{W2}: intersection=({X:F3},{Y:F3}), " +
                    "distW1={D1:F4}ft, distW2={D2:F4}ft",
                    wall1.Id, wall2.Id, intersection.X, intersection.Y,
                    distToWall1, distToWall2);
                return true;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "WallJoinReplicator: Failed to trim walls {W1} and {W2} to corner",
                    wall1.Id, wall2.Id);
            }

            return false;
        }

        /// <summary>
        /// Replaces the endpoint nearest to the target point with the target point.
        /// </summary>
        private static Line TrimLineToPoint(XYZ start, XYZ end, XYZ point)
        {
            if (start.DistanceTo(point) < end.DistanceTo(point))
                return Line.CreateBound(point, end);
            else
                return Line.CreateBound(start, point);
        }

        /// <summary>
        /// Returns a canonical ordered pair for use as a HashSet key.
        /// </summary>
        private static (long, long) OrderIds(ElementId a, ElementId b)
        {
            return a.Value < b.Value ? (a.Value, b.Value) : (b.Value, a.Value);
        }

        /// <summary>
        /// Toggles end-join state to force Revit to recalculate corner geometry.
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
                // Non-critical
            }
        }
    }
}
