using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Serilog;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Moves doors and windows from the original compound wall to the appropriate
    /// replacement wall after splitting. Must be called BEFORE deleting the original wall.
    /// </summary>
    public static class HostedElementTransfer
    {
        /// <summary>
        /// Finds all door and window FamilyInstances hosted on the specified wall.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="wall">The host wall to search for hosted elements.</param>
        /// <returns>List of FamilyInstances (doors/windows) hosted on the wall.</returns>
        public static List<FamilyInstance> GetHostedElements(Document doc, Wall wall)
        {
            if (doc == null || wall == null)
                return new List<FamilyInstance>();

            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Host?.Id == wall.Id)
                .Where(fi =>
                {
                    var catId = fi.Category?.Id;
                    if (catId == null) return false;
                    return catId == new ElementId(BuiltInCategory.OST_Doors)
                        || catId == new ElementId(BuiltInCategory.OST_Windows);
                })
                .ToList();
        }

        /// <summary>
        /// Determines which replacement wall should receive the hosted elements.
        /// If the user picked a specific layer index, returns that wall.
        /// Otherwise auto-picks: thickest layer, then structural, then first.
        /// </summary>
        /// <param name="replacements">List of replacement walls paired with their layer info.</param>
        /// <param name="userPickedLayerIndex">Optional layer index chosen by the user.</param>
        /// <returns>The target wall to host the transferred elements, or null if none found.</returns>
        public static Wall FindTargetWall(List<(Wall Wall, LayerInfo Layer)> replacements, int? userPickedLayerIndex)
        {
            if (replacements == null || replacements.Count == 0)
                return null;

            // If user explicitly picked a layer index, find that replacement
            if (userPickedLayerIndex.HasValue)
            {
                var match = replacements.FirstOrDefault(r => r.Layer.Index == userPickedLayerIndex.Value);
                if (match.Wall != null)
                    return match.Wall;
            }

            // Auto-pick: thickest layer first
            double maxThickness = replacements.Max(r => r.Layer.Thickness);
            var thickest = replacements
                .Where(r => r.Layer.Thickness >= maxThickness - 1e-9)
                .ToList();

            if (thickest.Count == 1)
                return thickest[0].Wall;

            // Tied on thickness: prefer structural layer
            var structural = thickest.Where(r => r.Layer.IsStructural).ToList();
            if (structural.Count > 0)
                return structural[0].Wall;

            // Still tied: return first
            return thickest[0].Wall;
        }

        /// <summary>
        /// Transfers hosted door/window elements from the original wall to the target replacement wall.
        /// Each element is recreated on the new host with matching parameters and flip state.
        /// Must be called within an active transaction and BEFORE deleting the original wall.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="hostedElements">The door/window instances to transfer.</param>
        /// <param name="targetWall">The replacement wall to host the elements.</param>
        /// <param name="level">The level for element placement.</param>
        public static void TransferElements(Document doc, List<FamilyInstance> hostedElements, Wall targetWall, Level level)
        {
            if (doc == null || hostedElements == null || targetWall == null || level == null)
                return;

            foreach (var original in hostedElements)
            {
                try
                {
                    TransferSingleElement(doc, original, targetWall, level);
                }
                catch (System.Exception ex)
                {
                    Log.Warning(ex, "Failed to transfer hosted element {ElementId} to wall {WallId}",
                        original.Id, targetWall.Id);
                }
            }
        }

        /// <summary>
        /// Transfers a single hosted element to the target wall, copying location, symbol,
        /// parameters, and flip state.
        /// </summary>
        private static void TransferSingleElement(Document doc, FamilyInstance original, Wall targetWall, Level level)
        {
            // Get the element's location point
            var locationPoint = original.Location as LocationPoint;
            if (locationPoint == null)
            {
                Log.Warning("Hosted element {ElementId} has no LocationPoint, skipping transfer", original.Id);
                return;
            }

            XYZ point = locationPoint.Point;

            // Get the FamilySymbol
            FamilySymbol symbol = original.Symbol;
            if (symbol == null)
            {
                Log.Warning("Hosted element {ElementId} has no FamilySymbol, skipping transfer", original.Id);
                return;
            }

            // Activate the symbol if needed
            if (!symbol.IsActive)
                symbol.Activate();

            // Create new instance on the target wall
            FamilyInstance newInstance = doc.Create.NewFamilyInstance(
                point, symbol, targetWall, level, StructuralType.NonStructural);

            if (newInstance == null)
            {
                Log.Warning("Failed to create new instance for element {ElementId}", original.Id);
                return;
            }

            // Copy sill height
            CopyParameter(original, newInstance, BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);

            // Copy head height
            CopyParameter(original, newInstance, BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM);

            // Copy comments
            CopyParameter(original, newInstance, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

            // Copy mark
            CopyParameter(original, newInstance, BuiltInParameter.ALL_MODEL_MARK);

            // Match flip state
            if (original.FacingFlipped != newInstance.FacingFlipped)
                newInstance.flipFacing();

            if (original.HandFlipped != newInstance.HandFlipped)
                newInstance.flipHand();
        }

        /// <summary>
        /// Copies a single parameter value from the original element to the new element,
        /// handling all storage types.
        /// </summary>
        private static void CopyParameter(FamilyInstance source, FamilyInstance target, BuiltInParameter paramId)
        {
            Parameter sourceParam = source.get_Parameter(paramId);
            Parameter targetParam = target.get_Parameter(paramId);

            if (sourceParam == null || targetParam == null || targetParam.IsReadOnly)
                return;

            switch (sourceParam.StorageType)
            {
                case StorageType.Integer:
                    targetParam.Set(sourceParam.AsInteger());
                    break;
                case StorageType.Double:
                    targetParam.Set(sourceParam.AsDouble());
                    break;
                case StorageType.String:
                    targetParam.Set(sourceParam.AsString());
                    break;
                case StorageType.ElementId:
                    targetParam.Set(sourceParam.AsElementId());
                    break;
            }
        }
    }
}
