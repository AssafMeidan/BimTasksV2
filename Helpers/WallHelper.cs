using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Static helper methods for wall geometry operations: retrieving wall types,
    /// calculating offsets, creating parallel walls, setting constraints, etc.
    /// </summary>
    public static class WallHelper
    {
        /// <summary>
        /// Retrieves a WallType by its name (case-insensitive).
        /// </summary>
        public static WallType GetWallTypeByName(Document doc, string wallTypeName)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(wt => wt.Name.Equals(wallTypeName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Retrieves the width of a WallType from its WALL_ATTR_WIDTH_PARAM parameter.
        /// </summary>
        public static double GetWallTypeWidth(WallType wallType)
        {
            Parameter widthParam = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
            return widthParam != null ? widthParam.AsDouble() : 0.0;
        }

        /// <summary>
        /// Gets the LocationCurve of a wall.
        /// </summary>
        public static LocationCurve GetWallLocationCurve(Wall wall)
        {
            return wall.Location as LocationCurve;
        }

        /// <summary>
        /// Gets the direction vector of a linear wall.
        /// Returns null if the wall's location curve is not a line.
        /// </summary>
        public static XYZ GetWallDirection(Wall wall)
        {
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve?.Curve is Line line)
            {
                return (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
            }
            return null;
        }

        /// <summary>
        /// Gets the width of a wall instance.
        /// </summary>
        public static double GetWallWidth(Wall wall)
        {
            return wall.Width;
        }

        /// <summary>
        /// Calculates the total offset distances for placing parallel walls on either side
        /// of an original wall, accounting for the thicknesses of original and new walls.
        /// </summary>
        /// <param name="originalWidth">Width of the original wall.</param>
        /// <param name="newWidth1">Width of the wall on the positive (exterior) side.</param>
        /// <param name="newWidth2">Width of the wall on the negative (interior) side.</param>
        /// <returns>Tuple of positive and negative offset distances.</returns>
        public static (double PositiveOffset, double NegativeOffset) CalculateOffsets(
            double originalWidth, double newWidth1, double newWidth2)
        {
            double positiveOffset = (originalWidth / 2.0) + (newWidth1 / 2.0);
            double negativeOffset = (originalWidth / 2.0) + (newWidth2 / 2.0);
            return (positiveOffset, negativeOffset);
        }

        /// <summary>
        /// Retrieves essential parameters of a wall needed for creating parallel cladding walls.
        /// </summary>
        public static (double OriginalWallWidth, ElementId BaseLevelId, double BaseOffset,
            ElementId TopConstraintId, double TopOffset, double WallHeight, bool IsUnconnected)
            GetWallParameters(Document doc, Wall wall)
        {
            double wallWidth = wall.Width;

            Parameter baseConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
            ElementId baseLevelId = baseConstraintParam.AsElementId();

            Parameter baseOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET);
            double baseOffset = baseOffsetParam.AsDouble();

            Parameter topConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
            ElementId topConstraintId = topConstraintParam.AsElementId();

            Parameter topOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
            double topOffset = topOffsetParam.AsDouble();

            double wallHeight = 0.0;
            bool isUnconnected = false;

            if (topConstraintId == ElementId.InvalidElementId)
            {
                isUnconnected = true;
                Parameter unconnectedHeightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeight = unconnectedHeightParam.AsDouble();
            }
            else
            {
                Level topLevel = doc.GetElement(topConstraintId) as Level;
                Level baseLevel = doc.GetElement(baseLevelId) as Level;

                if (topLevel == null || baseLevel == null)
                {
                    throw new InvalidOperationException("Invalid top or base level.");
                }

                double topLevelElevation = topLevel.Elevation;
                double baseLevelElevation = baseLevel.Elevation;
                wallHeight = (topLevelElevation + topOffset) - (baseLevelElevation + baseOffset);
            }

            return (wallWidth, baseLevelId, baseOffset, topConstraintId, topOffset, wallHeight, isUnconnected);
        }

        /// <summary>
        /// Creates a new parallel wall with specified parameters.
        /// </summary>
        public static Wall CreateParallelWall(
            Document doc, Curve curve, WallType wallType,
            ElementId baseLevelId, double wallHeight, double baseOffset, bool isFlipped)
        {
            Wall newWall = Wall.Create(
                doc,
                curve,
                wallType.Id,
                baseLevelId,
                wallHeight,
                baseOffset,
                isFlipped,
                false);

            return newWall;
        }

        /// <summary>
        /// Sets the top constraint and offset for a wall.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="wall">The wall to modify.</param>
        /// <param name="topConstraintId">The ElementId of the top constraint level.</param>
        /// <param name="topOffset">The top offset value.</param>
        /// <param name="isUnconnected">If true, sets the wall to unconnected height mode.</param>
        public static void SetTopConstraint(Document doc, Wall wall, ElementId topConstraintId, double topOffset, bool isUnconnected)
        {
            if (isUnconnected)
            {
                // Set Top Constraint to Unconnected
                Parameter topConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                if (topConstraintParam != null && !topConstraintParam.IsReadOnly)
                {
                    topConstraintParam.Set(ElementId.InvalidElementId);
                }

                // Set Unconnected Height (preserve height from creation)
                Parameter unconnectedHeightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (unconnectedHeightParam != null && !unconnectedHeightParam.IsReadOnly)
                {
                    unconnectedHeightParam.Set(wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble());
                }
            }
            else
            {
                // Set Top Constraint to same level as original
                Parameter topConstraintParam = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                if (topConstraintParam != null && !topConstraintParam.IsReadOnly)
                {
                    topConstraintParam.Set(topConstraintId);
                }

                // Set Top Offset
                Parameter topOffsetParam = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
                if (topOffsetParam != null && !topOffsetParam.IsReadOnly)
                {
                    topOffsetParam.Set(topOffset);
                }
            }
        }

        /// <summary>
        /// Sets the location line for a wall using the Location Line parameter.
        /// </summary>
        public static void SetWallLocationLine(Wall wall, WallLocationLineValue locationLine)
        {
            Parameter locationLineParam = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);

            if (locationLineParam != null && !locationLineParam.IsReadOnly)
            {
                locationLineParam.Set((int)locationLine);
            }
        }

        /// <summary>
        /// Copies additional parameters (Function, Structural Usage, Comments, Mark)
        /// from the original wall to the new wall.
        /// </summary>
        public static void CopyAdditionalParameters(Wall originalWall, Wall newWall)
        {
            CopyParameter(originalWall, newWall, BuiltInParameter.FUNCTION_PARAM);
            CopyParameter(originalWall, newWall, BuiltInParameter.WALL_STRUCTURAL_USAGE_PARAM);
            CopyParameter(originalWall, newWall, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            CopyParameter(originalWall, newWall, BuiltInParameter.ALL_MODEL_MARK);
        }

        /// <summary>
        /// Copies a specific parameter from one wall to another, handling all storage types.
        /// </summary>
        private static void CopyParameter(Wall originalWall, Wall newWall, BuiltInParameter paramId)
        {
            Parameter originalParam = originalWall.get_Parameter(paramId);
            Parameter newParam = newWall.get_Parameter(paramId);
            if (originalParam != null && newParam != null && !newParam.IsReadOnly)
            {
                switch (originalParam.StorageType)
                {
                    case StorageType.Integer:
                        newParam.Set(originalParam.AsInteger());
                        break;
                    case StorageType.Double:
                        newParam.Set(originalParam.AsDouble());
                        break;
                    case StorageType.String:
                        newParam.Set(originalParam.AsString());
                        break;
                    case StorageType.ElementId:
                        newParam.Set(originalParam.AsElementId());
                        break;
                }
            }
        }

        /// <summary>
        /// Joins a list of new walls to the original wall using JoinGeometryUtils.
        /// </summary>
        public static void JoinWalls(Document doc, Wall originalWall, List<Wall> newWalls)
        {
            foreach (Wall newWall in newWalls)
            {
                try
                {
                    if (!JoinGeometryUtils.AreElementsJoined(doc, originalWall, newWall))
                    {
                        JoinGeometryUtils.JoinGeometry(doc, originalWall, newWall);
                    }
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    // Walls may not be compatible for joining; skip silently
                }
            }
        }
    }
}
