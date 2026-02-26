using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Result of a single wall split operation.
    /// </summary>
    public class SplitResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public List<Wall> ReplacementWalls { get; set; } = new();
        /// <summary>
        /// Replacement walls paired with their layer info, preserving layer order
        /// (exterior→interior). Used by cross-join to prioritize matching types/positions.
        /// </summary>
        public List<(Wall Wall, LayerInfo Layer)> ReplacementPairs { get; set; } = new();
        public int TransferredElementCount { get; set; }
    }

    /// <summary>
    /// Orchestrates the wall split pipeline: analyzes compound layers, creates
    /// replacement walls, transfers hosted elements, deletes the original, and
    /// re-establishes geometry joins — all within a single rollback-capable TransactionGroup.
    /// </summary>
    public static class WallSplitterEngine
    {
        /// <summary>
        /// Splits a single compound wall into individual single-layer walls,
        /// one per non-membrane layer.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="wall">The compound wall to split.</param>
        /// <param name="layers">Pre-analyzed layer info (from CompoundLayerAnalyzer).</param>
        /// <param name="hostedElementTargetLayerIndex">
        /// Optional layer index that should receive hosted doors/windows.
        /// If null, the engine auto-picks the thickest/structural layer.
        /// </param>
        /// <returns>A SplitResult describing the outcome.</returns>
        public static SplitResult SplitWall(Document doc, Wall wall, List<LayerInfo> layers, int? hostedElementTargetLayerIndex)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (wall == null) throw new ArgumentNullException(nameof(wall));
            if (layers == null || layers.Count == 0)
                return new SplitResult { Success = false, Message = "No layers provided." };

            // Record joins BEFORE any transactions modify the wall
            JoinRecord joinRecord;
            try
            {
                joinRecord = WallJoinReplicator.RecordJoins(doc, wall);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[WallSplitterEngine] Failed to record joins for wall {WallId}, continuing without join replication", wall.Id);
                joinRecord = new JoinRecord();
            }

            // Read wall parameters and geometry BEFORE entering transactions
            var wallParams = WallHelper.GetWallParameters(doc, wall);
            XYZ orientation = wall.Orientation;
            Curve originalCurve = (wall.Location as LocationCurve)!.Curve;
            bool isFlipped = wall.Flipped;
            Level baseLevel = doc.GetElement(wallParams.BaseLevelId) as Level;

            var replacements = new List<(Wall Wall, LayerInfo Layer)>();
            int transferredCount = 0;

            using var txGroup = new TransactionGroup(doc, "Split Compound Wall");
            txGroup.Start();

            try
            {
                // === Transaction 1: Create Replacement Walls ===
                using (var tx1 = new Transaction(doc, "Create Replacement Walls"))
                {
                    var failOpts1 = tx1.GetFailureHandlingOptions();
                    failOpts1.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx1.SetFailureHandlingOptions(failOpts1);
                    tx1.Start();

                    foreach (var layer in layers)
                    {
                        // Resolve WallType if not already set
                        if (layer.ResolvedType == null)
                        {
                            WallTypeResolver.FindOrCreateType(doc, layer);
                        }

                        // Compute offset vector from wall orientation and layer center offset
                        XYZ offsetVector = orientation * layer.CenterOffset;

                        // Create translated curve at the layer's center offset
                        Curve translatedCurve = originalCurve.CreateTransformed(
                            Transform.CreateTranslation(offsetVector));

                        // Create the replacement wall
                        Wall newWall = WallHelper.CreateParallelWall(
                            doc,
                            translatedCurve,
                            layer.ResolvedType!,
                            wallParams.BaseLevelId,
                            wallParams.WallHeight,
                            wallParams.BaseOffset,
                            isFlipped);

                        // Set top constraint to match original
                        WallHelper.SetTopConstraint(doc, newWall,
                            wallParams.TopConstraintId, wallParams.TopOffset, wallParams.IsUnconnected);

                        // Force location line to wall centerline
                        WallHelper.SetWallLocationLine(newWall, WallLocationLineValue.WallCenterline);

                        // Copy additional parameters from original
                        WallHelper.CopyAdditionalParameters(wall, newWall);

                        replacements.Add((newWall, layer));
                    }

                    tx1.Commit();
                }

                // === Transaction 2: Transfer Hosted Elements ===
                using (var tx2 = new Transaction(doc, "Transfer Hosted Elements"))
                {
                    var failOpts2 = tx2.GetFailureHandlingOptions();
                    failOpts2.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx2.SetFailureHandlingOptions(failOpts2);
                    tx2.Start();

                    var hostedElements = HostedElementTransfer.GetHostedElements(doc, wall);
                    if (hostedElements.Count > 0)
                    {
                        Wall targetWall = HostedElementTransfer.FindTargetWall(replacements, hostedElementTargetLayerIndex);
                        if (targetWall != null && baseLevel != null)
                        {
                            HostedElementTransfer.TransferElements(doc, hostedElements, targetWall, baseLevel);
                            transferredCount = hostedElements.Count;
                        }
                    }

                    tx2.Commit();
                }

                // === Transaction 3: Delete Original Wall ===
                using (var tx3 = new Transaction(doc, "Delete Original Wall"))
                {
                    var failOpts3 = tx3.GetFailureHandlingOptions();
                    failOpts3.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx3.SetFailureHandlingOptions(failOpts3);
                    tx3.Start();

                    doc.Delete(wall.Id);

                    tx3.Commit();
                }

                // === Transaction 4: Re-establish Joins ===
                using (var tx4 = new Transaction(doc, "Re-establish Joins"))
                {
                    var failOpts4 = tx4.GetFailureHandlingOptions();
                    failOpts4.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx4.SetFailureHandlingOptions(failOpts4);
                    tx4.Start();

                    var replacementWalls = replacements.Select(r => r.Wall).ToList();
                    WallJoinReplicator.ReplicateJoins(doc, replacementWalls, joinRecord);

                    tx4.Commit();
                }

                txGroup.Assimilate();

                var result = new SplitResult
                {
                    Success = true,
                    Message = $"Split into {replacements.Count} walls. Transferred {transferredCount} hosted element(s).",
                    ReplacementWalls = replacements.Select(r => r.Wall).ToList(),
                    ReplacementPairs = replacements,
                    TransferredElementCount = transferredCount
                };

                Log.Information("[WallSplitterEngine] {Message}", result.Message);
                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[WallSplitterEngine] Failed to split wall {WallId}", wall.Id);
                txGroup.RollBack();
                return new SplitResult
                {
                    Success = false,
                    Message = $"Split failed: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Splits multiple walls that share the same wall type, re-analyzing layers per wall.
        /// After all individual splits, performs a cross-join pass to connect replacement
        /// walls from different original walls at corners/intersections.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="walls">The walls to split (should share the same compound wall type).</param>
        /// <param name="layers">Layer configuration (used as a template; re-analyzed per wall).</param>
        /// <param name="hostedElementTargetLayerIndex">Optional target layer index for hosted elements.</param>
        /// <returns>List of SplitResult, one per wall.</returns>
        public static List<SplitResult> SplitWalls(Document doc, List<Wall> walls, List<LayerInfo> layers, int? hostedElementTargetLayerIndex)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (walls == null) throw new ArgumentNullException(nameof(walls));

            var results = new List<SplitResult>();

            foreach (var wall in walls)
            {
                try
                {
                    // Re-analyze layers for each wall since offsets depend on per-instance properties
                    var wallLayers = CompoundLayerAnalyzer.AnalyzeLayers(wall);
                    if (wallLayers.Count == 0)
                    {
                        results.Add(new SplitResult
                        {
                            Success = false,
                            Message = $"Wall {wall.Id} has no splittable layers."
                        });
                        continue;
                    }

                    var result = SplitWall(doc, wall, wallLayers, hostedElementTargetLayerIndex);
                    results.Add(result);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[WallSplitterEngine] Unexpected error splitting wall {WallId}", wall.Id);
                    results.Add(new SplitResult
                    {
                        Success = false,
                        Message = $"Wall {wall.Id}: {ex.Message}"
                    });
                }
            }

            // Save corner fix data so the separate FixSplitCorners command can
            // trim/extend walls to form clean corners in a separate transaction.
            var successfulResults = results.Where(r => r.Success).ToList();
            if (successfulResults.Count > 0)
            {
                try
                {
                    string filePath = GetCornerFixDataPath();
                    SaveCornerFixData(filePath, successfulResults);
                    Log.Information("[WallSplitterEngine] Saved corner fix data to {FilePath}", filePath);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[WallSplitterEngine] Failed to save corner fix data");
                }
            }

            Log.Information("[WallSplitterEngine] Batch split complete: {Succeeded}/{Total} succeeded",
                results.Count(r => r.Success), results.Count);

            return results;
        }

        // =================================================================
        // Corner Fix: JSON persistence + trim/extend logic
        // =================================================================

        /// <summary>
        /// Path to the corner fix data file in the user's temp directory.
        /// </summary>
        public static string GetCornerFixDataPath()
        {
            return Path.Combine(Path.GetTempPath(), "BimTasksV2_SplitCornerData.json");
        }

        /// <summary>
        /// Serializable data for a single replacement wall.
        /// </summary>
        public class CornerFixWallEntry
        {
            public long WallId { get; set; }
            public int LayerIndex { get; set; }
            public string WallTypeName { get; set; } = "";
        }

        /// <summary>
        /// Serializable data for one split group (one original wall's replacements).
        /// </summary>
        public class CornerFixGroup
        {
            public List<CornerFixWallEntry> Replacements { get; set; } = new();
        }

        /// <summary>
        /// Root serializable data for the corner fix file.
        /// </summary>
        public class CornerFixData
        {
            public List<CornerFixGroup> SplitGroups { get; set; } = new();
        }

        /// <summary>
        /// Saves replacement wall data to a JSON file for the FixSplitCorners command.
        /// </summary>
        public static void SaveCornerFixData(string filePath, List<SplitResult> results)
        {
            var data = new CornerFixData();

            foreach (var result in results)
            {
                var group = new CornerFixGroup();
                foreach (var pair in result.ReplacementPairs)
                {
                    group.Replacements.Add(new CornerFixWallEntry
                    {
                        WallId = pair.Wall.Id.Value,
                        LayerIndex = pair.Layer.Index,
                        WallTypeName = pair.Wall.WallType.Name
                    });
                }
                data.SplitGroups.Add(group);
            }

            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(filePath, json);
        }

        /// <summary>
        /// Loads corner fix data from JSON and executes trim/extend on matching
        /// walls from different split groups. Runs in its own clean transaction.
        /// </summary>
        /// <returns>Summary message.</returns>
        public static string FixCorners(Document doc)
        {
            string filePath = GetCornerFixDataPath();
            if (!File.Exists(filePath))
                return "No corner fix data found. Run Split Wall first.";

            CornerFixData data;
            try
            {
                string json = File.ReadAllText(filePath);
                data = JsonSerializer.Deserialize<CornerFixData>(json);
            }
            catch (Exception ex)
            {
                return $"Failed to read corner fix data: {ex.Message}";
            }

            if (data == null || data.SplitGroups.Count == 0)
                return "Corner fix data is empty.";

            // Resolve wall IDs to actual Wall elements, skip any that no longer exist
            var resolvedGroups = new List<List<(Wall Wall, int LayerIndex, string TypeName)>>();
            foreach (var group in data.SplitGroups)
            {
                var resolved = new List<(Wall Wall, int LayerIndex, string TypeName)>();
                foreach (var entry in group.Replacements)
                {
                    var elem = doc.GetElement(new ElementId(entry.WallId)) as Wall;
                    if (elem != null)
                    {
                        resolved.Add((elem, entry.LayerIndex, entry.WallTypeName));
                    }
                    else
                    {
                        Log.Warning("[FixCorners] Wall {WallId} no longer exists, skipping", entry.WallId);
                    }
                }
                if (resolved.Count > 0)
                    resolvedGroups.Add(resolved);
            }

            if (resolvedGroups.Count < 2)
                return "Need at least 2 split groups to fix corners. Only found " + resolvedGroups.Count + ".";

            int trimmed = 0;

            using var tx = new Transaction(doc, "Fix Split Wall Corners");
            var failOpts = tx.GetFailureHandlingOptions();
            failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
            tx.SetFailureHandlingOptions(failOpts);
            tx.Start();

            try
            {
                var connected = new HashSet<(long, long)>();

                // For each pair of split groups (different original walls)
                for (int a = 0; a < resolvedGroups.Count; a++)
                {
                    for (int b = a + 1; b < resolvedGroups.Count; b++)
                    {
                        // Pass 1: Match by same WallType name
                        foreach (var wa in resolvedGroups[a])
                        {
                            foreach (var wb in resolvedGroups[b])
                            {
                                var key = OrderIds(wa.Wall.Id.Value, wb.Wall.Id.Value);
                                if (connected.Contains(key)) continue;

                                if (wa.TypeName == wb.TypeName)
                                {
                                    if (TrimWallsToCorner(wa.Wall, wb.Wall))
                                    {
                                        connected.Add(key);
                                        trimmed++;
                                    }
                                }
                            }
                        }

                        // Pass 2: Match by same layer index
                        foreach (var wa in resolvedGroups[a])
                        {
                            foreach (var wb in resolvedGroups[b])
                            {
                                var key = OrderIds(wa.Wall.Id.Value, wb.Wall.Id.Value);
                                if (connected.Contains(key)) continue;

                                if (wa.LayerIndex == wb.LayerIndex)
                                {
                                    if (TrimWallsToCorner(wa.Wall, wb.Wall))
                                    {
                                        connected.Add(key);
                                        trimmed++;
                                    }
                                }
                            }
                        }
                    }
                }

                // Toggle end joins on all walls
                foreach (var group in resolvedGroups)
                {
                    foreach (var entry in group)
                    {
                        ForceEndJoinRecalculation(entry.Wall, 0);
                        ForceEndJoinRecalculation(entry.Wall, 1);
                    }
                }

                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.RollBack();
                return $"Corner fix failed: {ex.Message}";
            }

            // Clean up the JSON file after successful fix
            try { File.Delete(filePath); } catch { }

            return $"Fixed {trimmed} corner connection(s) across {resolvedGroups.Count} split groups.";
        }

        // Tolerance for extending curves to detect near-intersections (in feet, ~10 cm)
        private const double ExtendTolerance = 1.0 / 30.48 * 10.0;

        /// <summary>
        /// Extends both wall curves, finds intersection, trims to meet at corner.
        /// </summary>
        private static bool TrimWallsToCorner(Wall wall1, Wall wall2)
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
                Log.Warning(ex, "[WallSplitterEngine] Failed to trim walls {W1} and {W2} to corner",
                    wall1.Id, wall2.Id);
            }
            return false;
        }

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

        private static Curve TrimToIntersection(Curve curve, XYZ point)
        {
            if (curve is Line)
            {
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                if (start.DistanceTo(point) < end.DistanceTo(point))
                    return Line.CreateBound(point, end);
                else
                    return Line.CreateBound(start, point);
            }
            return null;
        }

        private static void ForceEndJoinRecalculation(Wall wall, int end)
        {
            try
            {
                WallUtils.DisallowWallJoinAtEnd(wall, end);
                WallUtils.AllowWallJoinAtEnd(wall, end);
            }
            catch { }
        }

        private static (long, long) OrderIds(long a, long b)
        {
            return a < b ? (a, b) : (b, a);
        }
    }
}
