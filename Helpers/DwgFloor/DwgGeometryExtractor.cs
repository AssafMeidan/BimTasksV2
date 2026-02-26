using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers.DwgFloor
{
    /// <summary>
    /// Extracts geometry (faces and loose curves) from DWG ImportInstance elements.
    /// </summary>
    public static class DwgGeometryExtractor
    {
        public class ExtractionResult
        {
            public List<List<Curve>> FaceLoops { get; } = new();
            public List<Curve> LooseCurves { get; } = new();
            public int ErrorCount { get; set; }
        }

        /// <summary>
        /// Extract all horizontal face loops and loose curves from a DWG ImportInstance.
        /// When a view is provided, only geometry visible in that view is returned.
        /// </summary>
        public static ExtractionResult ExtractAll(ImportInstance importInstance, View? activeView = null)
        {
            var result = new ExtractionResult();
            var options = new Options { IncludeNonVisibleObjects = false };

            if (activeView != null)
                options.View = activeView;
            else
                options.DetailLevel = ViewDetailLevel.Fine;

            var geometry = importInstance.get_Geometry(options);
            if (geometry == null)
                return result;

            var transform = importInstance.GetTransform();
            ProcessGeometry(geometry, transform, 0, result);
            return result;
        }

        /// <summary>
        /// Try to find a horizontal face at/near a picked point from a DWG ImportInstance.
        /// Returns the outer edge loop curves, or null if not found.
        /// </summary>
        public static List<Curve>? FindFaceAtReference(
            ImportInstance importInstance,
            GeometryObject pickedGeom,
            Reference reference)
        {
            var transform = importInstance.GetTransform();

            if (pickedGeom is Face face)
                return ExtractFaceOuterLoop(face, transform);

            if (pickedGeom is Edge edge)
            {
                try
                {
                    var edgeCurve = edge.AsCurve();
                    return FindFaceContainingCurve(importInstance, edgeCurve, transform);
                }
                catch { return null; }
            }

            if (pickedGeom is Curve curve)
                return FindFaceContainingCurve(importInstance, curve, transform);

            return null;
        }

        /// <summary>
        /// Extract all faces and return the one whose centroid is nearest to the picked point.
        /// </summary>
        public static List<Curve>? FindNearestFace(ImportInstance importInstance, XYZ pickedPoint)
        {
            var extraction = ExtractAll(importInstance);
            if (extraction.FaceLoops.Count == 0)
                return null;

            var flatPicked = CurveFlattener.FlattenPoint(pickedPoint);
            List<Curve>? bestLoop = null;
            double bestDist = double.MaxValue;

            foreach (var faceLoop in extraction.FaceLoops)
            {
                var flatCurves = CurveFlattener.FlattenAll(faceLoop);
                if (flatCurves.Count < 3) continue;

                double sumX = 0, sumY = 0;
                int count = 0;
                foreach (var c in flatCurves)
                {
                    var pt = c.GetEndPoint(0);
                    sumX += pt.X; sumY += pt.Y; count++;
                }

                var centroid = count > 0 ? new XYZ(sumX / count, sumY / count, 0) : XYZ.Zero;
                double dist = flatPicked.DistanceTo(centroid);

                if (dist < bestDist)
                {
                    bestDist = dist;
                    bestLoop = faceLoop;
                }
            }

            return bestLoop;
        }

        private static List<Curve>? ExtractFaceOuterLoop(Face face, Transform transform)
        {
            try
            {
                if (!IsHorizontalFace(face)) return null;
                var outerLoop = face.GetEdgesAsCurveLoops().FirstOrDefault();
                if (outerLoop == null) return null;

                var curves = new List<Curve>();
                foreach (Curve curve in outerLoop)
                    curves.Add(curve.CreateTransformed(transform));
                return curves;
            }
            catch { return null; }
        }

        private static List<Curve>? FindFaceContainingCurve(
            ImportInstance importInstance, Curve targetCurve, Transform transform)
        {
            var options = new Options
            {
                IncludeNonVisibleObjects = false,
                DetailLevel = ViewDetailLevel.Fine
            };

            var geometry = importInstance.get_Geometry(options);
            if (geometry == null) return null;

            var targetMid = targetCurve.Evaluate(0.5, true);
            var targetStart = targetCurve.GetEndPoint(0);
            var targetEnd = targetCurve.GetEndPoint(1);

            return SearchForFaceContainingPoint(geometry, transform, targetMid, targetStart, targetEnd, 0);
        }

        private static List<Curve>? SearchForFaceContainingPoint(
            GeometryElement geometry, Transform transform,
            XYZ targetMid, XYZ targetStart, XYZ targetEnd, int depth)
        {
            foreach (var geomObj in geometry)
            {
                try
                {
                    if (geomObj is Solid solid && solid.Faces.Size > 0)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            try
                            {
                                if (!IsHorizontalFace(face)) continue;

                                foreach (var loop in face.GetEdgesAsCurveLoops())
                                {
                                    foreach (Curve curve in loop)
                                    {
                                        var mid = transform.OfPoint(curve.Evaluate(0.5, true));
                                        if (mid.DistanceTo(targetMid) < DwgFloorConfig.ClosureTolerance)
                                            return ExtractFaceOuterLoop(face, transform);

                                        var start = transform.OfPoint(curve.GetEndPoint(0));
                                        var end = transform.OfPoint(curve.GetEndPoint(1));
                                        if (start.DistanceTo(targetStart) < DwgFloorConfig.ClosureTolerance &&
                                            end.DistanceTo(targetEnd) < DwgFloorConfig.ClosureTolerance)
                                            return ExtractFaceOuterLoop(face, transform);
                                    }
                                }
                            }
                            catch { /* skip bad face */ }
                        }
                    }
                    else if (geomObj is GeometryInstance gi && depth < DwgFloorConfig.MaxRecursionDepth)
                    {
                        var childGeom = gi.GetSymbolGeometry();
                        if (childGeom != null)
                        {
                            var composed = transform.Multiply(gi.Transform);
                            var found = SearchForFaceContainingPoint(
                                childGeom, composed, targetMid, targetStart, targetEnd, depth + 1);
                            if (found != null) return found;
                        }
                    }
                }
                catch { /* skip bad geometry object */ }
            }
            return null;
        }

        private static void ProcessGeometry(
            GeometryElement geometry, Transform transform, int depth, ExtractionResult result)
        {
            foreach (var geomObj in geometry)
            {
                try
                {
                    switch (geomObj)
                    {
                        case Solid solid:
                            ProcessSolid(solid, transform, result);
                            break;

                        case GeometryInstance gi when depth < DwgFloorConfig.MaxRecursionDepth:
                            var childGeom = gi.GetSymbolGeometry();
                            if (childGeom != null)
                            {
                                var composed = transform.Multiply(gi.Transform);
                                ProcessGeometry(childGeom, composed, depth + 1, result);
                            }
                            break;

                        case Curve curve:
                            result.LooseCurves.Add(curve.CreateTransformed(transform));
                            break;

                        case PolyLine polyLine:
                            ProcessPolyLine(polyLine, transform, result);
                            break;
                    }
                }
                catch { result.ErrorCount++; }
            }
        }

        private static void ProcessSolid(Solid solid, Transform transform, ExtractionResult result)
        {
            try { if (solid.Faces.Size == 0) return; }
            catch { result.ErrorCount++; return; }

            foreach (Face face in solid.Faces)
            {
                try
                {
                    if (!IsHorizontalFace(face)) continue;
                    var outerLoop = face.GetEdgesAsCurveLoops().FirstOrDefault();
                    if (outerLoop == null) continue;

                    var curves = new List<Curve>();
                    foreach (Curve curve in outerLoop)
                        curves.Add(curve.CreateTransformed(transform));

                    if (curves.Count >= 3)
                        result.FaceLoops.Add(curves);
                }
                catch { result.ErrorCount++; }
            }
        }

        private static void ProcessPolyLine(PolyLine polyLine, Transform transform, ExtractionResult result)
        {
            var coords = polyLine.GetCoordinates();
            for (int i = 0; i < coords.Count - 1; i++)
            {
                try
                {
                    var p0 = transform.OfPoint(coords[i]);
                    var p1 = transform.OfPoint(coords[i + 1]);
                    if (p0.DistanceTo(p1) >= DwgFloorConfig.ShortCurveThreshold)
                        result.LooseCurves.Add(Line.CreateBound(p0, p1));
                }
                catch { result.ErrorCount++; }
            }
        }

        private static bool IsHorizontalFace(Face face)
        {
            try
            {
                var bb = face.GetBoundingBox();
                var uvMid = new UV(
                    (bb.Min.U + bb.Max.U) / 2.0,
                    (bb.Min.V + bb.Max.V) / 2.0);
                var normal = face.ComputeNormal(uvMid);
                return Math.Abs(normal.Z) > DwgFloorConfig.HorizontalNormalThreshold;
            }
            catch { return false; }
        }
    }
}
