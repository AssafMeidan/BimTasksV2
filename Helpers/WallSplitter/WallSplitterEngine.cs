using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

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

            // Cross-join pass: connect replacement walls from different original walls
            // at corners and intersections (needed because original neighbor IDs become
            // invalid after the neighbor is also split).
            // Uses prioritized joining: same WallType first, then same layer position,
            // then remaining pairs.
            var allPairs = results
                .Where(r => r.Success)
                .SelectMany(r => r.ReplacementPairs)
                .ToList();

            if (allPairs.Count > 1)
            {
                try
                {
                    using var txCrossJoin = new Transaction(doc, "Cross-Join Split Walls");
                    var failOpts = txCrossJoin.GetFailureHandlingOptions();
                    failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    txCrossJoin.SetFailureHandlingOptions(failOpts);
                    txCrossJoin.Start();

                    WallJoinReplicator.CrossJoinReplacements(doc, allPairs);

                    txCrossJoin.Commit();
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "[WallSplitterEngine] Cross-join pass failed (non-critical)");
                }
            }

            Log.Information("[WallSplitterEngine] Batch split complete: {Succeeded}/{Total} succeeded",
                results.Count(r => r.Success), results.Count);

            return results;
        }
    }
}
