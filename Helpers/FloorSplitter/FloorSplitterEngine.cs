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
    /// Holds slope data extracted from the original floor.
    /// </summary>
    public class FloorSlopeData
    {
        /// <summary>Slope arrow line (null if no slope arrow).</summary>
        public Line? SlopeArrow { get; set; }

        /// <summary>Raw slope value for Floor.Create (from ROOF_SLOPE or computed via Atan).</summary>
        public double SlopeAngle { get; set; }

        /// <summary>True if the floor has shape-edited vertices.</summary>
        public bool HasShapeEditing { get; set; }

        /// <summary>
        /// Shape-edited vertex positions with their Z offsets from the base plane.
        /// </summary>
        public List<(XYZ Position, double Offset)> VertexOffsets { get; set; } = new();
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
            var level = doc.GetElement(levelId) as Level;
            double levelZ = level?.Elevation ?? 0;
            double heightOffset = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)?.AsDouble() ?? 0.0;

            // Extract the floor boundary from its sketch
            var curveLoops = GetFloorBoundary(doc, floor);
            if (curveLoops == null || curveLoops.Count == 0)
                return new FloorSplitResult { Success = false, Message = "Could not extract floor boundary." };

            // Record openings hosted on this floor BEFORE any modifications
            // Flatten Z to level elevation so openings work on sloped floors
            var openingBoundaries = RecordOpenings(doc, floor, levelZ + heightOffset);

            // Record slope data (slope arrow or shape editing)
            var slopeData = RecordSlope(doc, floor);

            var replacements = new List<(Floor Floor, FloorLayerInfo Layer)>();

            using var txGroup = new TransactionGroup(doc, "Split Compound Floor");
            txGroup.Start();

            try
            {
                // === Transaction 1: Create Replacement Floors WITH slope ===
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

                        // Create with slope arrow if original had one
                        var newFloor = Floor.Create(doc, curveLoops, layer.ResolvedType!.Id, levelId,
                            true, slopeData.SlopeArrow, slopeData.SlopeAngle);
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

                // === Transaction 1.5: Apply Shape Editing (if original had modified vertices) ===
                if (slopeData.HasShapeEditing && slopeData.VertexOffsets.Count > 0)
                {
                    using (var txShape = new Transaction(doc, "Apply Shape Editing"))
                    {
                        var failOpts = txShape.GetFailureHandlingOptions();
                        failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                        txShape.SetFailureHandlingOptions(failOpts);
                        txShape.Start();

                        foreach (var replacement in replacements)
                        {
                            ApplyShapeEditing(replacement.Floor, slopeData.VertexOffsets);
                        }

                        txShape.Commit();
                    }
                }

                // === Transaction 2: Recreate Openings (with flattened Z curves) ===
                int openingsCreated = 0;
                int openingsFailed = 0;
                if (openingBoundaries.Count > 0)
                {
                    using (var tx2 = new Transaction(doc, "Recreate Floor Openings"))
                    {
                        var failOpts = tx2.GetFailureHandlingOptions();
                        failOpts.SetFailuresPreprocessor(new SuppressWarningsPreprocessor());
                        tx2.SetFailureHandlingOptions(failOpts);
                        tx2.Start();

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
                                    openingsFailed++;
                                    Log.Warning(ex, "[FloorSplitterEngine] Failed to create opening on floor {FloorId}",
                                        replacement.Floor.Id);
                                }
                            }
                        }

                        Log.Information("[FloorSplitterEngine] Openings: {Created} created, {Failed} failed across {Floors} floors",
                            openingsCreated, openingsFailed, replacements.Count);

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
                string openingMsg = openingsCreated > 0 ? $" {openingsCreated} opening(s) transferred." : "";
                if (openingsFailed > 0)
                    openingMsg += $" {openingsFailed} opening(s) failed.";
                string slopeMsg = slopeData.SlopeArrow != null ? " Slope preserved." :
                    slopeData.HasShapeEditing ? " Shape editing preserved." : "";
                var result = new FloorSplitResult
                {
                    Success = true,
                    Message = $"Split into {replacements.Count} floors.{openingMsg}{slopeMsg}",
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
        /// Records slope data from the original floor: slope arrow and/or shape editing.
        /// Must be called BEFORE any transactions modify the floor.
        /// </summary>
        private static FloorSlopeData RecordSlope(Document doc, Floor floor)
        {
            var data = new FloorSlopeData();

            try
            {
                // === Method 1: Find slope arrow via dependent elements ===
                var dependentIds = floor.GetDependentElements(null);
                foreach (var id in dependentIds)
                {
                    var elem = doc.GetElement(id);
                    if (elem == null) continue;

                    // Slope arrows have SLOPE_START_HEIGHT parameter
                    var startHeightParam = elem.get_Parameter(BuiltInParameter.SLOPE_START_HEIGHT);
                    if (startHeightParam == null) continue;

                    // Found a slope arrow element — get its geometry
                    if (elem.Location is LocationCurve locCurve && locCurve.Curve is Line arrowLine)
                    {
                        data.SlopeArrow = arrowLine;

                        // Get the slope value from the element (for Floor.Create)
                        var slopeParam = elem.get_Parameter(BuiltInParameter.ROOF_SLOPE);
                        if (slopeParam != null)
                        {
                            data.SlopeAngle = slopeParam.AsDouble();
                        }
                        else
                        {
                            // Compute slope from start/end heights
                            double startH = startHeightParam.AsDouble();
                            var endHeightParam = elem.get_Parameter(BuiltInParameter.SLOPE_END_HEIGHT);
                            double endH = endHeightParam?.AsDouble() ?? startH;
                            double length = arrowLine.Length;
                            if (length > 1e-9)
                                data.SlopeAngle = Math.Atan((endH - startH) / length);
                        }

                        Log.Information("[FloorSplitterEngine] Found slope arrow: slopeValue={Slope:F6} on floor {FloorId}",
                            data.SlopeAngle, floor.Id);
                        break;
                    }
                }

                // === Method 2: Check SlabShapeEditor for shape-edited vertices ===
                var editor = floor.GetSlabShapeEditor();
                if (editor != null && editor.IsEnabled)
                {
                    data.HasShapeEditing = true;

                    // Compute base plane Z (where vertices would be if floor were flat)
                    var level = doc.GetElement(floor.LevelId) as Level;
                    double heightOff = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)?.AsDouble() ?? 0.0;
                    double basePlaneZ = (level?.Elevation ?? 0) + heightOff;

                    foreach (SlabShapeVertex vertex in editor.SlabShapeVertices)
                    {
                        double offset = vertex.Position.Z - basePlaneZ;
                        data.VertexOffsets.Add((vertex.Position, offset));
                    }

                    Log.Information("[FloorSplitterEngine] Found {Count} shape-edited vertices on floor {FloorId}",
                        data.VertexOffsets.Count, floor.Id);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Failed to extract slope data from floor {FloorId}", floor.Id);
            }

            return data;
        }

        /// <summary>
        /// Applies shape editing (vertex offsets) to a replacement floor,
        /// matching vertices by XY proximity to the original floor's vertices.
        /// </summary>
        private static void ApplyShapeEditing(Floor newFloor, List<(XYZ Position, double Offset)> originalOffsets)
        {
            const double xyTolerance = 0.01; // ~3mm tolerance for vertex matching

            try
            {
                var editor = newFloor.GetSlabShapeEditor();
                if (editor == null) return;

                editor.Enable();

                foreach (SlabShapeVertex newVertex in editor.SlabShapeVertices)
                {
                    // Find the matching original vertex by XY position
                    foreach (var (origPos, origOffset) in originalOffsets)
                    {
                        double dx = Math.Abs(newVertex.Position.X - origPos.X);
                        double dy = Math.Abs(newVertex.Position.Y - origPos.Y);

                        if (dx < xyTolerance && dy < xyTolerance && Math.Abs(origOffset) > 1e-9)
                        {
                            editor.ModifySubElement(newVertex, origOffset);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Failed to apply shape editing to floor {FloorId}", newFloor.Id);
            }
        }

        /// <summary>
        /// Records the boundary curves of all openings hosted on a floor.
        /// Flattens all Z coordinates to flatZ so openings can be recreated on sloped floors.
        /// Must be called BEFORE the original floor is deleted.
        /// </summary>
        private static List<CurveArray> RecordOpenings(Document doc, Floor floor, double flatZ)
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
                            // Rectangular opening — build CurveArray with flattened Z
                            var pts = opening.BoundaryRect;
                            if (pts != null && pts.Count >= 2)
                            {
                                var min = pts[0];
                                var max = pts[1];
                                var curveArray = new CurveArray();
                                curveArray.Append(Line.CreateBound(new XYZ(min.X, min.Y, flatZ), new XYZ(max.X, min.Y, flatZ)));
                                curveArray.Append(Line.CreateBound(new XYZ(max.X, min.Y, flatZ), new XYZ(max.X, max.Y, flatZ)));
                                curveArray.Append(Line.CreateBound(new XYZ(max.X, max.Y, flatZ), new XYZ(min.X, max.Y, flatZ)));
                                curveArray.Append(Line.CreateBound(new XYZ(min.X, max.Y, flatZ), new XYZ(min.X, min.Y, flatZ)));
                                boundaries.Add(curveArray);
                            }
                        }
                        else
                        {
                            // Sketch-based opening — flatten curves to consistent Z
                            var boundaryCurves = opening.BoundaryCurves;
                            if (boundaryCurves != null && boundaryCurves.Size > 0)
                            {
                                var curveArray = new CurveArray();
                                foreach (Curve curve in boundaryCurves)
                                {
                                    curveArray.Append(FlattenCurve(curve, flatZ));
                                }
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
                    Log.Information("[FloorSplitterEngine] Recorded {Count} opening(s) from floor {FloorId} (flatZ={FlatZ:F4})",
                        boundaries.Count, floor.Id, flatZ);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "[FloorSplitterEngine] Failed to query openings for floor {FloorId}", floor.Id);
            }

            return boundaries;
        }

        /// <summary>
        /// Projects a curve to a constant Z elevation while preserving its XY shape.
        /// </summary>
        private static Curve FlattenCurve(Curve curve, double z)
        {
            if (curve is Line line)
            {
                var p0 = line.GetEndPoint(0);
                var p1 = line.GetEndPoint(1);
                return Line.CreateBound(new XYZ(p0.X, p0.Y, z), new XYZ(p1.X, p1.Y, z));
            }
            else if (curve is Arc arc)
            {
                var p0 = arc.GetEndPoint(0);
                var p1 = arc.GetEndPoint(1);
                var mid = arc.Evaluate(0.5, true);
                return Arc.Create(
                    new XYZ(p0.X, p0.Y, z),
                    new XYZ(p1.X, p1.Y, z),
                    new XYZ(mid.X, mid.Y, z));
            }
            else
            {
                // For other curve types, try tessellation fallback
                // Just return as-is — most openings are lines/arcs
                return curve;
            }
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
