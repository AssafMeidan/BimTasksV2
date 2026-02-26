using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers.DwgFloor
{
    /// <summary>
    /// Builds CurveLoops from curve lists and creates Floor elements.
    /// </summary>
    public static class FloorCreator
    {
        public class FloorCreationResult
        {
            public int Created { get; set; }
            public int Failed { get; set; }
            public List<string> Errors { get; } = new();
        }

        /// <summary>
        /// Build a CurveLoop from a list of curves, closing micro-gaps.
        /// </summary>
        public static CurveLoop? BuildCurveLoop(List<Curve> curves)
        {
            if (curves.Count < 3)
                return null;

            var flatCurves = CurveFlattener.FlattenAll(curves);
            if (flatCurves.Count < 3)
                return null;

            var adjusted = CloseMicroGaps(flatCurves);
            if (adjusted == null || adjusted.Count < 3)
                return null;

            try
            {
                var loop = new CurveLoop();
                foreach (var curve in adjusted)
                    loop.Append(curve);
                return loop;
            }
            catch { return null; }
        }

        /// <summary>
        /// Compute area of a curve loop using the shoelace formula on XY projection.
        /// </summary>
        public static double ComputeAreaSqft(CurveLoop loop)
        {
            var points = new List<XYZ>();
            foreach (Curve curve in loop)
                points.Add(curve.GetEndPoint(0));

            if (points.Count < 3) return 0.0;

            double area = 0.0;
            for (int i = 0; i < points.Count; i++)
            {
                var current = points[i];
                var next = points[(i + 1) % points.Count];
                area += current.X * next.Y - next.X * current.Y;
            }
            return Math.Abs(area) / 2.0;
        }

        /// <summary>
        /// Compute the centroid of a curve loop on XY plane.
        /// </summary>
        public static XYZ ComputeCentroid(CurveLoop loop)
        {
            double sumX = 0, sumY = 0;
            int count = 0;
            foreach (Curve curve in loop)
            {
                var pt = curve.GetEndPoint(0);
                sumX += pt.X;
                sumY += pt.Y;
                count++;
            }
            return count > 0 ? new XYZ(sumX / count, sumY / count, 0) : XYZ.Zero;
        }

        /// <summary>
        /// Create a single floor from a curve loop.
        /// </summary>
        public static Floor? CreateFloor(
            Document doc, CurveLoop loop, ElementId floorTypeId, ElementId levelId)
        {
            try
            {
                var curveLoops = new List<CurveLoop> { loop };
                return Floor.Create(doc, curveLoops, floorTypeId, levelId, true, null, 0);
            }
            catch { return null; }
        }

        /// <summary>
        /// Create floors from multiple curve loops in a single transaction.
        /// </summary>
        public static FloorCreationResult CreateFloors(
            Document doc, List<CurveLoop> loops, ElementId floorTypeId, ElementId levelId)
        {
            var result = new FloorCreationResult();

            using var trans = new Transaction(doc, "Import DWG Floors");
            trans.Start();

            for (int i = 0; i < loops.Count; i++)
            {
                try
                {
                    var floor = CreateFloor(doc, loops[i], floorTypeId, levelId);
                    if (floor != null)
                        result.Created++;
                    else
                    {
                        result.Failed++;
                        result.Errors.Add($"Floor {i + 1}: creation returned null");
                    }
                }
                catch (Exception ex)
                {
                    result.Failed++;
                    result.Errors.Add($"Floor {i + 1}: {ex.Message}");
                }
            }

            if (result.Created > 0)
                trans.Commit();
            else
                trans.RollBack();

            return result;
        }

        /// <summary>
        /// Close micro-gaps between consecutive curve endpoints.
        /// </summary>
        private static List<Curve>? CloseMicroGaps(List<Curve> curves)
        {
            if (curves.Count < 3) return null;

            // First pass: check all gaps are within tolerance
            for (int i = 0; i < curves.Count; i++)
            {
                int next = (i + 1) % curves.Count;
                double gap = curves[i].GetEndPoint(1).DistanceTo(curves[next].GetEndPoint(0));
                if (gap > DwgFloorConfig.ClosureTolerance)
                    return null;
            }

            // Second pass: rebuild curves with adjusted endpoints
            var final = new List<Curve>(curves.Count);
            for (int i = 0; i < curves.Count; i++)
            {
                int next = (i + 1) % curves.Count;
                var curve = curves[i];

                var start = (i == 0) ? curve.GetEndPoint(0) : final[i - 1].GetEndPoint(1);
                var end = curves[next].GetEndPoint(0);

                // Last curve connects back to first curve's start
                if (i == curves.Count - 1)
                    end = final[0].GetEndPoint(0);

                try
                {
                    Curve? newCurve;
                    if (curve is Line)
                    {
                        if (start.DistanceTo(end) < DwgFloorConfig.ShortCurveThreshold) continue;
                        newCurve = Line.CreateBound(start, end);
                    }
                    else if (curve is Arc arc)
                    {
                        double param = (arc.GetEndParameter(0) + arc.GetEndParameter(1)) / 2.0;
                        var mid = CurveFlattener.FlattenPoint(arc.Evaluate(param, false));
                        if (start.DistanceTo(end) < DwgFloorConfig.ShortCurveThreshold) continue;
                        try { newCurve = Arc.Create(start, end, mid); }
                        catch { newCurve = Line.CreateBound(start, end); }
                    }
                    else
                    {
                        if (start.DistanceTo(end) < DwgFloorConfig.ShortCurveThreshold) continue;
                        newCurve = Line.CreateBound(start, end);
                    }

                    final.Add(newCurve);
                }
                catch { return null; }
            }

            return final.Count >= 3 ? final : null;
        }
    }
}
