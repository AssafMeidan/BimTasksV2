using Autodesk.Revit.DB;
using Serilog;
using System.Collections.Generic;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Records the join state of an original wall before splitting,
    /// then replicates adjacent-layer joins onto the replacement walls.
    /// Cross-wall corner connections are handled by the Fix Corners panel.
    /// </summary>
    public class JoinRecord
    {
        public List<ElementId> GeometryJoinedElements { get; set; } = new();
        public bool AllowJoinAtStart { get; set; }
        public bool AllowJoinAtEnd { get; set; }
    }

    public static class WallJoinReplicator
    {
        /// <summary>
        /// Captures the current join state of a wall before it is split/deleted.
        /// </summary>
        public static JoinRecord RecordJoins(Document doc, Wall wall)
        {
            var record = new JoinRecord();

            var joinedIds = JoinGeometryUtils.GetJoinedElements(doc, wall);
            foreach (var id in joinedIds)
            {
                record.GeometryJoinedElements.Add(id);
            }

            record.AllowJoinAtStart = WallUtils.IsWallJoinAllowedAtEnd(wall, 0);
            record.AllowJoinAtEnd = WallUtils.IsWallJoinAllowedAtEnd(wall, 1);

            return record;
        }

        /// <summary>
        /// Joins adjacent replacement layers of the SAME original wall to each other.
        /// These physically overlap so JoinGeometry works correctly.
        /// Cross-wall corner connections are handled separately by FixSplitCorners.
        /// </summary>
        public static void ReplicateJoins(Document doc, List<Wall> replacementWalls, JoinRecord record)
        {
            if (replacementWalls == null || replacementWalls.Count == 0)
                return;

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
        }
    }
}
