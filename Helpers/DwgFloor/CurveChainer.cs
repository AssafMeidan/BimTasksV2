using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace BimTasksV2.Helpers.DwgFloor
{
    /// <summary>
    /// Chains loose curves into closed loops by proximity matching.
    /// </summary>
    public static class CurveChainer
    {
        /// <summary>
        /// Chain a collection of loose curves into closed loops.
        /// </summary>
        public static List<List<Curve>> ChainCurves(List<Curve> curves)
        {
            var result = new List<List<Curve>>();

            var flatCurves = CurveFlattener.FlattenAll(curves);
            if (flatCurves.Count == 0)
                return result;

            var deduped = DeduplicateCurves(flatCurves);
            if (deduped.Count < 3)
                return result;

            var used = new bool[deduped.Count];

            for (int startIdx = 0; startIdx < deduped.Count; startIdx++)
            {
                if (used[startIdx]) continue;

                var loop = TryBuildLoop(deduped, used, startIdx);
                if (loop != null && loop.Count >= 3)
                    result.Add(loop);
            }

            return result;
        }

        private static List<Curve>? TryBuildLoop(List<Curve> curves, bool[] used, int startIdx)
        {
            var loop = new List<Curve>();
            used[startIdx] = true;
            loop.Add(curves[startIdx]);

            var firstStart = curves[startIdx].GetEndPoint(0);
            var currentEnd = curves[startIdx].GetEndPoint(1);

            int maxIterations = curves.Count;
            for (int iter = 0; iter < maxIterations; iter++)
            {
                if (loop.Count >= 3 && currentEnd.DistanceTo(firstStart) < DwgFloorConfig.ChainTolerance)
                    return loop;

                int bestIdx = -1;
                double bestDist = DwgFloorConfig.ChainTolerance;
                bool bestReversed = false;

                for (int i = 0; i < curves.Count; i++)
                {
                    if (used[i]) continue;

                    double distToStart = currentEnd.DistanceTo(curves[i].GetEndPoint(0));
                    double distToEnd = currentEnd.DistanceTo(curves[i].GetEndPoint(1));

                    if (distToStart < bestDist)
                    {
                        bestDist = distToStart;
                        bestIdx = i;
                        bestReversed = false;
                    }
                    if (distToEnd < bestDist)
                    {
                        bestDist = distToEnd;
                        bestIdx = i;
                        bestReversed = true;
                    }
                }

                if (bestIdx < 0) break;

                used[bestIdx] = true;
                var nextCurve = bestReversed ? ReverseCurve(curves[bestIdx]) : curves[bestIdx];
                if (nextCurve == null) break;

                loop.Add(nextCurve);
                currentEnd = nextCurve.GetEndPoint(1);
            }

            if (loop.Count >= 3 && currentEnd.DistanceTo(firstStart) < DwgFloorConfig.ChainTolerance)
                return loop;

            return null;
        }

        private static Curve? ReverseCurve(Curve curve)
        {
            try { return curve.CreateReversed(); }
            catch
            {
                try
                {
                    if (curve is Line)
                        return Line.CreateBound(curve.GetEndPoint(1), curve.GetEndPoint(0));
                    if (curve is Arc arc)
                    {
                        double param = (arc.GetEndParameter(0) + arc.GetEndParameter(1)) / 2.0;
                        var mid = arc.Evaluate(param, false);
                        return Arc.Create(curve.GetEndPoint(1), curve.GetEndPoint(0), mid);
                    }
                }
                catch { }
                return null;
            }
        }

        private static List<Curve> DeduplicateCurves(List<Curve> curves)
        {
            var result = new List<Curve>();

            for (int i = 0; i < curves.Count; i++)
            {
                bool isDup = false;
                var midI = curves[i].Evaluate(0.5, true);
                var startI = curves[i].GetEndPoint(0);
                var endI = curves[i].GetEndPoint(1);

                for (int j = 0; j < result.Count; j++)
                {
                    var midJ = result[j].Evaluate(0.5, true);
                    var startJ = result[j].GetEndPoint(0);
                    var endJ = result[j].GetEndPoint(1);

                    bool midMatch = midI.DistanceTo(midJ) < DwgFloorConfig.DeduplicationTolerance;
                    bool endpointsMatch =
                        (startI.DistanceTo(startJ) < DwgFloorConfig.DeduplicationTolerance &&
                         endI.DistanceTo(endJ) < DwgFloorConfig.DeduplicationTolerance) ||
                        (startI.DistanceTo(endJ) < DwgFloorConfig.DeduplicationTolerance &&
                         endI.DistanceTo(startJ) < DwgFloorConfig.DeduplicationTolerance);

                    if (midMatch && endpointsMatch)
                    {
                        isDup = true;
                        break;
                    }
                }

                if (!isDup)
                    result.Add(curves[i]);
            }

            return result;
        }
    }
}
