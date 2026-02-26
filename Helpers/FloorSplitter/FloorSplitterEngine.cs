using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.FloorSplitter
{
    /// <summary>
    /// Result of a single floor split operation.
    /// </summary>
    public class FloorSplitResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public List<Floor> ReplacementFloors { get; set; } = new();
        public List<(Floor Floor, FloorLayerInfo Layer)> ReplacementPairs { get; set; } = new();
    }

    /// <summary>
    /// Orchestrates the floor split pipeline: creates replacement single-layer floors
    /// at the correct vertical offsets, joins adjacent layers, and deletes the original.
    /// </summary>
    public static class FloorSplitterEngine
    {
        /// <summary>
        /// Splits a single compound floor into individual single-layer floors.
        /// </summary>
        public static FloorSplitResult SplitFloor(Document doc, Floor floor, List<FloorLayerInfo> layers)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (floor == null) throw new ArgumentNullException(nameof(floor));
            if (layers == null || layers.Count == 0)
                return new FloorSplitResult { Success = false, Message = "No layers provided." };

            // Read floor parameters before any modifications
            var levelId = floor.LevelId;
            double heightOffset = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)?.AsDouble() ?? 0.0;

            // Extract the floor boundary from its sketch
            var curveLoops = GetFloorBoundary(doc, floor);
            if (curveLoops == null || curveLoops.Count == 0)
                return new FloorSplitResult { Success = false, Message = "Could not extract floor boundary." };

            // Record openings hosted on this floor BEFORE any modifications
            var openingBoundaries = RecordOpenings(doc, floor);

            var replacements = new List<(Floor Floor, FloorLayerInfo Layer)>();

            using var txGroup = new TransactionGroup(doc, "Split Compound Floor");
            txGroup.Start();

            try
            {
                // === Transaction 1: Create Replacement Floors ===
                using (var tx1 = new Transaction(doc, "Create Replacement Floors"))
                {
                    var failOpts = tx1.GetFailureHandlingOptions();
                    failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx1.SetFailureHandlingOptions(failOpts);
                    tx1.Start();

                    foreach (var layer in layers)
                    {
                        // Resolve FloorType if not already set
                        if (layer.ResolvedType == null)
                            FloorTypeResolver.FindOrCreateType(doc, layer);

                        // Create the replacement floor with same boundary
                        var newFloor = Floor.Create(doc, curveLoops, layer.ResolvedType!.Id, levelId, true, null, 0);
                        if (newFloor == null)
                        {
                            Log.Warning("[FloorSplitterEngine] Failed to create floor for layer {Index}", layer.Index);
                            continue;
                        }

                        // Set the height offset: original offset + layer's vertical offset
                        var heightParam = newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                        if (heightParam != null && !heightParam.IsReadOnly)
                            heightParam.Set(heightOffset + layer.TopOffset);

                        // Copy additional parameters
                        CopyFloorParameters(floor, newFloor);

                        replacements.Add((newFloor, layer));
                    }

                    tx1.Commit();
                }

                // === Transaction 2: Recreate Openings on Replacement Floors ===
                if (openingBoundaries.Count > 0)
                {
                    using (var tx2 = new Transaction(doc, "Recreate Floor Openings"))
                    {
                        var failOpts = tx2.GetFailureHandlingOptions();
                        failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                        tx2.SetFailureHandlingOptions(failOpts);
                        tx2.Start();

                        int openingsCreated = 0;
                        foreach (var replacement in replacements)
                        {
                            foreach (var boundary in openingBoundaries)
                            {
                                try
                                {
                                    doc.Create.NewOpening(replacement.Floor, boundary, false);
                                    openingsCreated++;
                                }
                                catch (Exception ex)
                                {
                                    Log.Warning(ex, "[FloorSplitterEngine] Failed to create opening on floor {FloorId}",
                                        replacement.Floor.Id);
                                }
                            }
                        }

                        Log.Information("[FloorSplitterEngine] Recreated {Count} openings across {Floors} replacement floors",
                            openingsCreated, replacements.Count);

                        tx2.Commit();
                    }
                }

                // === Transaction 3: Delete Original Floor ===
                using (var tx3 = new Transaction(doc, "Delete Original Floor"))
                {
                    var failOpts = tx3.GetFailureHandlingOptions();
                    failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx3.SetFailureHandlingOptions(failOpts);
                    tx3.Start();

                    doc.Delete(floor.Id);

                    tx3.Commit();
                }

                // === Transaction 4: Join Adjacent Layers ===
                using (var tx4 = new Transaction(doc, "Join Adjacent Floor Layers"))
                {
                    var failOpts = tx4.GetFailureHandlingOptions();
                    failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                    tx4.SetFailureHandlingOptions(failOpts);
                    tx4.Start();

                    for (int i = 0; i < replacements.Count - 1; i++)
                    {
                        try
                        {
                            JoinGeometryUtils.JoinGeometry(doc, replacements[i].Floor, replacements[i + 1].Floor);
                        }
                        catch
                        {
                            Log.Warning("[FloorSplitterEngine] Failed to join floor layers {A} and {B}",
                                replacements[i].Floor.Id, replacements[i + 1].Floor.Id);
                        }
                    }

                    tx4.Commit();
                }

                txGroup.Assimilate();

                int openingCount = openingBoundaries.Count;
                string openingMsg = openingCount > 0 ? $" {openingCount} opening(s) transferred." : "";
                var result = new FloorSplitResult
                {
                    Success = true,
                    Message = $"Split into {replacements.Count} floors.{openingMsg}",
                    ReplacementFloors = replacements.Select(r => r.Floor).ToList(),
                    ReplacementPairs = replacements
                };

                Log.Information("[FloorSplitterEngine] {Message}", result.Message);
                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[FloorSplitterEngine] Failed to split floor {FloorId}", floor.Id);
                txGroup.RollBack();
                return new FloorSplitResult
                {
                    Success = false,
                    Message = $"Split failed: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Records the boundary curves of all openings hosted on a floor.
        /// Must be called BEFORE the original floor is deleted.
        /// </summary>
        private static List<CurveArray> RecordOpenings(Document doc, Floor floor)
        {
            var boundaries = new List<CurveArray>();

            try
            {
                // Find all floor openings hosted on this floor
                var openings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FloorOpening)
                    .OfClass(typeof(Opening))
                    .Cast<Opening>()
                    .Where(o => o.Host?.Id == floor.Id)
                    .ToList();

                foreach (var opening in openings)
                {
                    try
                    {
                        if (opening.IsRectBoundary)
                        {
                            // Rectangular opening — build CurveArray from BoundaryRect points
                            var pts = opening.BoundaryRect;
                            if (pts != null && pts.Count >= 2)
                            {
                                var min = pts[0];
                                var max = pts[1];
                                var curveArray = new CurveArray();
                                curveArray.Append(Line.CreateBound(new XYZ(min.X, min.Y, min.Z), new XYZ(max.X, min.Y, min.Z)));
                                curveArray.Append(Line.CreateBound(new XYZ(max.X, min.Y, min.Z), new XYZ(max.X, max.Y, min.Z)));
                                curveArray.Append(Line.CreateBound(new XYZ(max.X, max.Y, min.Z), new XYZ(min.X, max.Y, min.Z)));
                                curveArray.Append(Line.CreateBound(new XYZ(min.X, max.Y, min.Z), new XYZ(min.X, min.Y, min.Z)));
                                boundaries.Add(curveArray);
                            }
                        }
                        else
                        {
                            // Sketch-based opening — copy boundary curves
                            var boundaryCurves = opening.BoundaryCurves;
                            if (boundaryCurves != null && boundaryCurves.Size > 0)
                            {
                                var curveArray = new CurveArray();
                                foreach (Curve curve in boundaryCurves)
                                    curveArray.Append(curve);
                                boundaries.Add(curveArray);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "[FloorSplitterEngine] Failed to record opening {OpeningId}", opening.Id);
                    }
                }

                if (boundaries.Count > 0)
                    Log.Information("[FloorSplitterEngine] Recorded {Count} opening(s) from floor {FloorId}",
                        boundaries.Count, floor.Id);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Failed to query openings for floor {FloorId}", floor.Id);
            }

            return boundaries;
        }

        /// <summary>
        /// Extracts the boundary CurveLoops from a floor's sketch.
        /// </summary>
        private static IList<CurveLoop> GetFloorBoundary(Document doc, Floor floor)
        {
            try
            {
                var sketch = doc.GetElement(floor.SketchId) as Sketch;
                if (sketch != null)
                {
                    var profile = sketch.Profile;
                    var curveLoops = new List<CurveLoop>();
                    foreach (CurveArray curveArray in profile)
                    {
                        var loop = new CurveLoop();
                        foreach (Curve curve in curveArray)
                            loop.Append(curve);
                        curveLoops.Add(loop);
                    }
                    return curveLoops;
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Failed to get floor sketch, trying geometry approach");
            }

            // Fallback: extract from top face geometry
            try
            {
                var opt = new Options { ComputeReferences = false, DetailLevel = ViewDetailLevel.Fine };
                var geomElem = floor.get_Geometry(opt);
                if (geomElem == null) return null;

                foreach (var geomObj in geomElem)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        // Find the topmost horizontal face
                        Face topFace = null;
                        double maxZ = double.MinValue;

                        foreach (Face face in solid.Faces)
                        {
                            if (face is PlanarFace planar && Math.Abs(planar.FaceNormal.Z - 1.0) < 0.01)
                            {
                                var bb = face.GetBoundingBox();
                                double z = planar.Origin.Z;
                                if (z > maxZ)
                                {
                                    maxZ = z;
                                    topFace = face;
                                }
                            }
                        }

                        if (topFace != null)
                            return topFace.GetEdgesAsCurveLoops();
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Geometry fallback failed for floor {FloorId}", floor.Id);
            }

            return null;
        }

        /// <summary>
        /// Copies common parameters from the original floor to the replacement.
        /// </summary>
        private static void CopyFloorParameters(Floor original, Floor newFloor)
        {
            CopyParameter(original, newFloor, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            CopyParameter(original, newFloor, BuiltInParameter.ALL_MODEL_MARK);
            CopyParameter(original, newFloor, BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
        }

        private static void CopyParameter(Floor source, Floor target, BuiltInParameter paramId)
        {
            try
            {
                var srcParam = source.get_Parameter(paramId);
                var tgtParam = target.get_Parameter(paramId);
                if (srcParam == null || tgtParam == null || tgtParam.IsReadOnly) return;

                switch (srcParam.StorageType)
                {
                    case StorageType.Integer: tgtParam.Set(srcParam.AsInteger()); break;
                    case StorageType.Double: tgtParam.Set(srcParam.AsDouble()); break;
                    case StorageType.String: tgtParam.Set(srcParam.AsString()); break;
                    case StorageType.ElementId: tgtParam.Set(srcParam.AsElementId()); break;
                }
            }
            catch { }
        }
    }
}
