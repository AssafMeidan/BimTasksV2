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
    }
}
