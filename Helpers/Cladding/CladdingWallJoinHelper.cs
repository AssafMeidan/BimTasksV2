using Autodesk.Revit.DB;
using BimTasksV2.Models;
using System;
using System.Collections.Generic;

namespace BimTasksV2.Helpers.Cladding
{
    /// <summary>
    /// Handles joining/connecting cladding wall ends to each other.
    /// Detects proximity and angles between cladding walls and trims
    /// or extends them to form clean intersections.
    /// </summary>
    public static class CladdingWallJoinHelper
    {
        private const double CmToFeet = 1.0 / 30.48;
        private static readonly double _touchDistance = 10.0 * CmToFeet;
        private static readonly double _tolerance = 0.8;

        /// <summary>
        /// Connects cladding walls by detecting their proximity, trimming to
        /// intersections, and enabling end joins.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="claddingWalls">The list of cladding wall members to connect.</param>
        public static void ConnectCladdingWalls(Document doc, List<MyWallMember> claddingWalls)
        {
            for (int i = 0; i < claddingWalls.Count; i++)
            {
                for (int j = i + 1; j < claddingWalls.Count; j++)
                {
                    Wall wall1 = claddingWalls[i].MyWall;
                    Wall wall2 = claddingWalls[j].MyWall;

                    if (AreWallsIntersecting(wall1, wall2))
                    {
                        TrimWallsToIntersection(doc, wall1, wall2);
                    }
                }
            }

            using (Transaction tx = new Transaction(doc, "Join Wall Ends"))
            {
                tx.Start();
                foreach (var wall in claddingWalls)
                {
                    AllowEndJoin(wall.MyWall);
                }
                tx.Commit();
            }
        }

        /// <summary>
        /// Determines if two walls intersect by extending their curves slightly
        /// and checking for overlap.
        /// </summary>
        private static bool AreWallsIntersecting(Wall wall1, Wall wall2)
        {
            LocationCurve locCurve1 = wall1.Location as LocationCurve;
            LocationCurve locCurve2 = wall2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                return false;

            Curve extendedCurve1 = ExtendCurve(locCurve1.Curve, _tolerance);
            Curve extendedCurve2 = ExtendCurve(locCurve2.Curve, _tolerance);

            IntersectionResultArray resultArray;
            SetComparisonResult result = extendedCurve1.Intersect(extendedCurve2, out resultArray);

            return result == SetComparisonResult.Overlap;
        }

        /// <summary>
        /// Trims two walls to the point of intersection.
        /// </summary>
        private static void TrimWallsToIntersection(Document doc, Wall wall1, Wall wall2)
        {
            LocationCurve locCurve1 = wall1.Location as LocationCurve;
            LocationCurve locCurve2 = wall2.Location as LocationCurve;

            if (locCurve1 == null || locCurve2 == null)
                throw new InvalidOperationException("One or both walls do not have valid LocationCurves.");

            Curve extendedCurve1 = ExtendCurve(locCurve1.Curve, _tolerance);
            Curve extendedCurve2 = ExtendCurve(locCurve2.Curve, _tolerance);

            IntersectionResultArray resultArray;
            SetComparisonResult result = extendedCurve1.Intersect(extendedCurve2, out resultArray);

            if (result == SetComparisonResult.Overlap && resultArray != null && resultArray.Size > 0)
            {
                XYZ intersectionPoint = resultArray.get_Item(0).XYZPoint;

                using (Transaction tx = new Transaction(doc, "Trim Walls to Intersection"))
                {
                    tx.Start();
                    locCurve1.Curve = TrimCurveToIntersection(locCurve1.Curve, intersectionPoint);
                    locCurve2.Curve = TrimCurveToIntersection(locCurve2.Curve, intersectionPoint);
                    tx.Commit();
                }
            }
        }

        /// <summary>
        /// Extends a line curve by a specified length in both directions.
        /// Non-line curves are returned unchanged.
        /// </summary>
        private static Curve ExtendCurve(Curve curve, double extensionLength)
        {
            if (curve is Line line)
            {
                XYZ startPoint = line.GetEndPoint(0);
                XYZ endPoint = line.GetEndPoint(1);
                XYZ direction = (endPoint - startPoint).Normalize();

                XYZ newStartPoint = startPoint - (direction * extensionLength);
                XYZ newEndPoint = endPoint + (direction * extensionLength);

                return Line.CreateBound(newStartPoint, newEndPoint);
            }

            return curve;
        }

        /// <summary>
        /// Trims a line curve to a specified intersection point. The farther endpoint
        /// from the intersection is kept, and the nearer endpoint is replaced.
        /// </summary>
        private static Curve TrimCurveToIntersection(Curve curve, XYZ intersectionPoint)
        {
            if (curve is Line line)
            {
                XYZ startPoint = line.GetEndPoint(0);
                XYZ endPoint = line.GetEndPoint(1);

                return (startPoint.DistanceTo(intersectionPoint) > endPoint.DistanceTo(intersectionPoint))
                    ? Line.CreateBound(startPoint, intersectionPoint)
                    : Line.CreateBound(intersectionPoint, endPoint);
            }

            return curve;
        }

        /// <summary>
        /// Allows the ends of a wall to join with other geometry.
        /// </summary>
        public static void AllowEndJoin(Wall wall)
        {
            WallUtils.AllowWallJoinAtEnd(wall, 0);
            WallUtils.AllowWallJoinAtEnd(wall, 1);
        }

        /// <summary>
        /// Disallows the ends of a wall from joining with other geometry.
        /// </summary>
        public static void DisallowEndJoin(Wall wall)
        {
            WallUtils.DisallowWallJoinAtEnd(wall, 0);
            WallUtils.DisallowWallJoinAtEnd(wall, 1);
        }
    }
}
