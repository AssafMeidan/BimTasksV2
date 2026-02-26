using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers.DwgFloor
{
    /// <summary>
    /// Flattens 3D curves to Z=0 for floor boundary creation.
    /// </summary>
    public static class CurveFlattener
    {
        /// <summary>
        /// Flatten a curve to Z=0. Returns null if the resulting curve is too short.
        /// </summary>
        public static Curve? Flatten(Curve curve)
        {
            try
            {
                return curve switch
                {
                    Line line => FlattenLine(line),
                    Arc arc => FlattenArc(arc),
                    _ => null // Complex curves handled by FlattenAll with tessellation
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Flatten a list of curves to Z=0, filtering out nulls and short curves.
        /// Complex curves (NurbSpline, HermiteSpline, Ellipse) are tessellated into line segments.
        /// </summary>
        public static List<Curve> FlattenAll(IEnumerable<Curve> curves)
        {
            var result = new List<Curve>();
            foreach (var curve in curves)
            {
                if (curve is Line or Arc)
                {
                    var flat = Flatten(curve);
                    if (flat != null)
                        result.Add(flat);
                }
                else
                {
                    var segments = TessellateToSegments(curve);
                    result.AddRange(segments);
                }
            }
            return result;
        }

        public static XYZ FlattenPoint(XYZ point) => new(point.X, point.Y, 0.0);

        private static Curve? FlattenLine(Line line)
        {
            var p0 = FlattenPoint(line.GetEndPoint(0));
            var p1 = FlattenPoint(line.GetEndPoint(1));
            if (p0.DistanceTo(p1) < DwgFloorConfig.ShortCurveThreshold)
                return null;
            return Line.CreateBound(p0, p1);
        }

        private static Curve? FlattenArc(Arc arc)
        {
            var p0 = FlattenPoint(arc.GetEndPoint(0));
            var p1 = FlattenPoint(arc.GetEndPoint(1));

            double param = (arc.GetEndParameter(0) + arc.GetEndParameter(1)) / 2.0;
            var mid = FlattenPoint(arc.Evaluate(param, false));

            if (p0.DistanceTo(p1) < DwgFloorConfig.ShortCurveThreshold)
                return null;
            if (p0.DistanceTo(mid) < DwgFloorConfig.ShortCurveThreshold ||
                p1.DistanceTo(mid) < DwgFloorConfig.ShortCurveThreshold)
                return Line.CreateBound(p0, p1);

            // Check collinearity
            var v1 = (mid - p0).Normalize();
            var v2 = (p1 - p0).Normalize();
            double cross = Math.Abs(v1.X * v2.Y - v1.Y * v2.X);
            if (cross < 1e-6)
                return Line.CreateBound(p0, p1);

            try
            {
                return Arc.Create(p0, p1, mid);
            }
            catch
            {
                return Line.CreateBound(p0, p1);
            }
        }

        /// <summary>
        /// Tessellate a complex curve into multiple line segments at Z=0.
        /// </summary>
        public static List<Curve> TessellateToSegments(Curve curve)
        {
            var result = new List<Curve>();
            try
            {
                double startParam = curve.GetEndParameter(0);
                double endParam = curve.GetEndParameter(1);
                int segments = DwgFloorConfig.TessellationSegments;

                for (int i = 0; i < segments; i++)
                {
                    double t0 = startParam + (endParam - startParam) * i / segments;
                    double t1 = startParam + (endParam - startParam) * (i + 1) / segments;

                    var p0 = FlattenPoint(curve.Evaluate(t0, false));
                    var p1 = FlattenPoint(curve.Evaluate(t1, false));

                    if (p0.DistanceTo(p1) >= DwgFloorConfig.ShortCurveThreshold)
                        result.Add(Line.CreateBound(p0, p1));
                }
            }
            catch
            {
                // Fallback to built-in tessellate
                try
                {
                    var pts = curve.Tessellate();
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        var p0 = FlattenPoint(pts[i]);
                        var p1 = FlattenPoint(pts[i + 1]);
                        if (p0.DistanceTo(p1) >= DwgFloorConfig.ShortCurveThreshold)
                            result.Add(Line.CreateBound(p0, p1));
                    }
                }
                catch { /* Skip entirely */ }
            }
            return result;
        }
    }
}
