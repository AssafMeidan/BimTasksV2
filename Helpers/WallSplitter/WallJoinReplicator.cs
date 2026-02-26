using Autodesk.Revit.DB;
using Serilog;
using System.Collections.Generic;

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
        /// Joins all replacement walls from different original walls to each other.
        /// Called after batch splitting to establish cross-wall joins.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="allReplacementWalls">All replacement walls from all split operations.</param>
        public static void CrossJoinReplacements(Document doc, List<Wall> allReplacementWalls)
        {
            if (allReplacementWalls == null || allReplacementWalls.Count < 2)
                return;

            // Try geometry-joining every pair (Revit will reject non-overlapping pairs)
            for (int i = 0; i < allReplacementWalls.Count; i++)
            {
                for (int j = i + 1; j < allReplacementWalls.Count; j++)
                {
                    try
                    {
                        if (!JoinGeometryUtils.AreElementsJoined(doc, allReplacementWalls[i], allReplacementWalls[j]))
                        {
                            JoinGeometryUtils.JoinGeometry(doc, allReplacementWalls[i], allReplacementWalls[j]);
                        }
                    }
                    catch
                    {
                        // Expected for non-overlapping walls — skip silently
                    }
                }
            }

            // Toggle end joins on all walls to force Revit corner cleanup
            foreach (var wall in allReplacementWalls)
            {
                ForceEndJoinRecalculation(wall, 0);
                ForceEndJoinRecalculation(wall, 1);
            }
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
