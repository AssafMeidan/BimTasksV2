using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Static helper methods for selecting and validating walls in Revit.
    /// </summary>
    public static class WallSelectionHelper
    {
        /// <summary>
        /// Retrieves walls from the current selection in the active document.
        /// Validates that each selected element is a non-curtain wall with a valid LocationCurve.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument.</param>
        /// <returns>A list of valid Wall objects from the current selection.</returns>
        public static List<Wall> GetSelectedWalls(UIDocument uiDoc)
        {
            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedIds = selection.GetElementIds();
            var selectedWalls = new List<Wall>();

            if (selectedIds.Count < 1)
            {
                return selectedWalls;
            }

            foreach (ElementId id in selectedIds)
            {
                Wall wall = uiDoc.Document.GetElement(id) as Wall;
                if (wall == null)
                    continue;

                if (IsValidWall(wall))
                {
                    selectedWalls.Add(wall);
                }
            }

            return selectedWalls;
        }

        /// <summary>
        /// Prompts the user to select walls from the current selection or shows
        /// an error if no valid walls are selected.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument.</param>
        /// <returns>A list of valid selected Wall objects.</returns>
        public static List<Wall> SelectWalls(UIDocument uiDoc)
        {
            var selectedWalls = GetSelectedWalls(uiDoc);

            if (selectedWalls.Count == 0)
            {
                TaskDialog.Show("Selection Error", "No valid walls were selected. Please select at least one wall.");
            }

            return selectedWalls;
        }

        /// <summary>
        /// Retrieves walls from the current selection, including non-linear walls.
        /// Only filters out non-wall elements and curtain walls.
        /// </summary>
        /// <param name="uiDoc">The active UIDocument.</param>
        /// <returns>A list of Wall objects (including non-linear walls).</returns>
        public static List<Wall> GetSelectedWallsIncludingNonLinear(UIDocument uiDoc)
        {
            Selection selection = uiDoc.Selection;
            ICollection<ElementId> selectedIds = selection.GetElementIds();
            var selectedWalls = new List<Wall>();

            foreach (ElementId id in selectedIds)
            {
                Wall wall = uiDoc.Document.GetElement(id) as Wall;
                if (wall != null && wall.CurtainGrid == null)
                {
                    selectedWalls.Add(wall);
                }
            }

            return selectedWalls;
        }

        /// <summary>
        /// Validates that a wall is suitable for operations (not a curtain wall,
        /// has a valid LocationCurve, and is linear).
        /// </summary>
        /// <param name="wall">The wall to validate.</param>
        /// <returns>True if the wall is valid for processing.</returns>
        public static bool IsValidWall(Wall wall)
        {
            if (wall == null)
                return false;

            // Exclude curtain walls
            if (wall.CurtainGrid != null)
                return false;

            // Must have a valid LocationCurve
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve == null)
                return false;

            // Must be a linear wall
            if (!(locCurve.Curve is Line))
                return false;

            return true;
        }

        /// <summary>
        /// Validates a list of walls and returns only the valid ones.
        /// </summary>
        /// <param name="walls">The list of walls to validate.</param>
        /// <returns>A filtered list of valid walls.</returns>
        public static List<Wall> ValidateWallSelection(IEnumerable<Wall> walls)
        {
            return walls.Where(IsValidWall).ToList();
        }
    }
}
