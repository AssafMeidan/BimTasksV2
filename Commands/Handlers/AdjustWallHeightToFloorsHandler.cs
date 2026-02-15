using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using BimTasksV2.Commands.Infrastructure;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Adjusts wall heights to match floor levels above and below.
    /// Uses ReferenceIntersector to find the nearest floors above and below each wall,
    /// then adjusts the wall's base offset (to sit on the floor below) and top
    /// offset/unconnected height (to reach the floor above).
    /// </summary>
    public class AdjustWallHeightToFloorsHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            ICollection<ElementId> selectedIds = sel.GetElementIds();
            if (selectedIds.Count == 0)
            {
                TaskDialog.Show("Selection", "Please select one or more walls.");
                return;
            }

            // Filter selected elements to walls only
            var selectedWalls = new List<Wall>();
            int invalidCount = 0;

            foreach (ElementId id in selectedIds)
            {
                Element element = doc.GetElement(id);
                if (element is Wall wall)
                {
                    selectedWalls.Add(wall);
                }
                else
                {
                    invalidCount++;
                }
            }

            if (selectedWalls.Count == 0)
            {
                TaskDialog.Show("Selection", "No walls were selected.");
                return;
            }

            if (invalidCount > 0)
            {
                TaskDialog.Show("Warning", "Some selected elements are not walls and will be ignored.");
            }

            // Get a 3D view for ray casting
            View3D view3D = Get3DView(doc);
            if (view3D == null)
            {
                TaskDialog.Show("Error", "No 3D view available for ray casting.");
                return;
            }

            // 1 meter offset for ray casting (in internal units)
            double offset = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Meters);

            // Set up the ReferenceIntersector for floors
            var floorFilter = new ElementClassFilter(typeof(Floor));
            var intersector = new ReferenceIntersector(floorFilter, FindReferenceTarget.Face, view3D);

            using (Transaction tx = new Transaction(doc, "Adjust Walls Height"))
            {
                tx.Start();

                foreach (Wall wall in selectedWalls)
                {
                    AdjustWallHeight(doc, wall, intersector, offset);
                }

                tx.Commit();
            }

            TaskDialog.Show("Completed", $"Adjusted {selectedWalls.Count} wall(s) to floor levels.");
        }

        private void AdjustWallHeight(Document doc, Wall wall, ReferenceIntersector intersector, double offset)
        {
            LocationCurve locCurve = wall.Location as LocationCurve;
            if (locCurve == null) return;

            Curve curve = locCurve.Curve;

            // Get wall elevation parameters
            double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
            Level baseLevel = doc.GetElement(
                wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()) as Level;

            if (baseLevel == null) return;

            Level topLevel = null;
            double topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).AsDouble();
            ElementId topConstraintId = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId();
            if (topConstraintId != ElementId.InvalidElementId)
            {
                topLevel = doc.GetElement(topConstraintId) as Level;
            }

            double baseElevation = baseLevel.Elevation + baseOffset;
            double topElevation;
            if (topLevel != null)
            {
                topElevation = topLevel.Elevation + topOffset;
            }
            else
            {
                double unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
                topElevation = baseElevation + unconnectedHeight;
            }

            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);

            XYZ baseStartPoint = new XYZ(startPoint.X, startPoint.Y, baseElevation);
            XYZ baseEndPoint = new XYZ(endPoint.X, endPoint.Y, baseElevation);
            XYZ topStartPoint = new XYZ(startPoint.X, startPoint.Y, topElevation);
            XYZ topEndPoint = new XYZ(endPoint.X, endPoint.Y, topElevation);

            // Find the floor below the wall's base
            double floorBelowTopElevation = double.MinValue;
            foreach (XYZ point in new[] { baseStartPoint, baseEndPoint })
            {
                XYZ adjustedPoint = point + XYZ.BasisZ * offset;
                XYZ directionDown = -XYZ.BasisZ;
                ReferenceWithContext refWithContext = intersector.FindNearest(adjustedPoint, directionDown);
                if (refWithContext != null)
                {
                    Reference reference = refWithContext.GetReference();
                    Floor floor = doc.GetElement(reference.ElementId) as Floor;
                    if (floor != null)
                    {
                        double floorTop = GetFloorTopElevation(floor);
                        if (floorTop > floorBelowTopElevation)
                        {
                            floorBelowTopElevation = floorTop;
                        }
                    }
                }
            }

            // Find the floor above the wall's top
            double floorAboveBottomElevation = double.MaxValue;
            foreach (XYZ point in new[] { topStartPoint, topEndPoint })
            {
                XYZ adjustedPoint = point - XYZ.BasisZ * offset;
                XYZ directionUp = XYZ.BasisZ;
                ReferenceWithContext refWithContext = intersector.FindNearest(adjustedPoint, directionUp);
                if (refWithContext != null)
                {
                    Reference reference = refWithContext.GetReference();
                    Floor floor = doc.GetElement(reference.ElementId) as Floor;
                    if (floor != null)
                    {
                        double floorBottom = GetFloorBottomElevation(floor);
                        if (floorBottom < floorAboveBottomElevation)
                        {
                            floorAboveBottomElevation = floorBottom;
                        }
                    }
                }
            }

            // Adjust base offset
            if (floorBelowTopElevation > double.MinValue)
            {
                double newBaseOffset = floorBelowTopElevation - baseLevel.Elevation;
                wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(newBaseOffset);
                baseElevation = floorBelowTopElevation;
            }

            // Adjust top offset or unconnected height
            if (floorAboveBottomElevation < double.MaxValue)
            {
                if (topLevel != null)
                {
                    double newTopOffset = floorAboveBottomElevation - topLevel.Elevation;
                    wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(newTopOffset);
                }
                else
                {
                    double newUnconnectedHeight = floorAboveBottomElevation - baseElevation;
                    wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(newUnconnectedHeight);
                }
            }
        }

        private double GetFloorTopElevation(Floor floor)
        {
            BoundingBoxXYZ bbox = floor.get_BoundingBox(null);
            return bbox != null ? bbox.Max.Z : 0;
        }

        private double GetFloorBottomElevation(Floor floor)
        {
            BoundingBoxXYZ bbox = floor.get_BoundingBox(null);
            return bbox != null ? bbox.Min.Z : 0;
        }

        private View3D Get3DView(Document doc)
        {
            // Try active view first
            if (doc.ActiveView is View3D active3DView && !active3DView.IsTemplate)
            {
                return active3DView;
            }

            // Find the default {3D} view
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && v.Name.Equals("{3D}"));
        }
    }
}
