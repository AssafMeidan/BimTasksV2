using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Data describing a single compound layer with its offset from the wall's location line.
    /// </summary>
    public class LayerInfo
    {
        public int Index { get; set; }
        public double Thickness { get; set; }
        public ElementId MaterialId { get; set; }
        public string MaterialName { get; set; } = "";
        public MaterialFunctionAssignment Function { get; set; }
        public bool IsStructural => Function == MaterialFunctionAssignment.Structure;

        /// <summary>
        /// Signed offset from the original wall's location line to this layer's center.
        /// Positive = toward exterior (wall.Orientation direction).
        /// </summary>
        public double CenterOffset { get; set; }

        /// <summary>
        /// The WallType to use for the replacement wall (resolved or auto-created).
        /// </summary>
        public WallType? ResolvedType { get; set; }

        /// <summary>
        /// True if the WallType was auto-created (doesn't exist in project yet).
        /// </summary>
        public bool IsAutoCreatedType { get; set; }
    }

    /// <summary>
    /// Reads CompoundStructure layers and computes each layer's center offset
    /// from the wall's location line.
    /// </summary>
    public static class CompoundLayerAnalyzer
    {
        /// <summary>
        /// Analyzes a compound wall and returns info for each non-membrane layer.
        /// </summary>
        /// <param name="wall">The compound wall to analyze.</param>
        /// <returns>List of LayerInfo objects, one per non-zero-thickness layer.</returns>
        public static List<LayerInfo> AnalyzeLayers(Wall wall)
        {
            var result = new List<LayerInfo>();
            var doc = wall.Document;
            var wallType = wall.WallType;
            var cs = wallType.GetCompoundStructure();

            if (cs == null)
                return result;

            var layers = cs.GetLayers();
            if (layers.Count <= 1)
                return result; // Not compound

            double totalWidth = wallType.Width;

            // Compute location line offset from exterior face
            double locationLineFromExterior = GetLocationLineOffsetFromExterior(wall, cs, totalWidth);

            // Running distance from exterior face
            double runningDistance = 0.0;

            for (int i = 0; i < layers.Count; i++)
            {
                var layer = layers[i];
                double thickness = layer.Width;

                // Skip zero-thickness membrane layers
                if (thickness < 1e-9)
                    continue;

                double layerCenterFromExterior = runningDistance + thickness / 2.0;

                // Offset from location line: positive = toward exterior = wall.Orientation direction
                double centerOffset = locationLineFromExterior - layerCenterFromExterior;

                // Account for wall flip
                if (wall.Flipped)
                    centerOffset = -centerOffset;

                string materialName = "";
                if (layer.MaterialId != ElementId.InvalidElementId)
                {
                    var mat = doc.GetElement(layer.MaterialId) as Material;
                    if (mat != null)
                        materialName = mat.Name;
                }

                result.Add(new LayerInfo
                {
                    Index = i,
                    Thickness = thickness,
                    MaterialId = layer.MaterialId,
                    MaterialName = materialName,
                    Function = layer.Function,
                    CenterOffset = centerOffset
                });

                runningDistance += thickness;
            }

            return result;
        }

        /// <summary>
        /// Determines how far the wall's location line sits from the exterior face,
        /// based on the WALL_KEY_REF_PARAM value.
        /// </summary>
        private static double GetLocationLineOffsetFromExterior(Wall wall, CompoundStructure cs, double totalWidth)
        {
            Parameter locLineParam = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
            int locLineValue = locLineParam?.AsInteger() ?? 0;
            var locLine = (WallLocationLineValue)locLineValue;

            switch (locLine)
            {
                case WallLocationLineValue.WallCenterline:
                    return totalWidth / 2.0;

                case WallLocationLineValue.CoreCenterline:
                {
                    double extFinish = GetExteriorFinishThickness(cs);
                    double coreWidth = GetCoreWidth(cs);
                    return extFinish + coreWidth / 2.0;
                }

                case WallLocationLineValue.FinishFaceExterior:
                    return 0.0;

                case WallLocationLineValue.FinishFaceInterior:
                    return totalWidth;

                case WallLocationLineValue.CoreExterior:
                    return GetExteriorFinishThickness(cs);

                case WallLocationLineValue.CoreInterior:
                    return GetExteriorFinishThickness(cs) + GetCoreWidth(cs);

                default:
                    return totalWidth / 2.0;
            }
        }

        /// <summary>
        /// Total thickness of layers outside the structural core.
        /// </summary>
        private static double GetExteriorFinishThickness(CompoundStructure cs)
        {
            var layers = cs.GetLayers();
            double sum = 0.0;
            int coreStart = cs.GetFirstCoreLayerIndex();
            for (int i = 0; i < coreStart && i < layers.Count; i++)
            {
                sum += layers[i].Width;
            }
            return sum;
        }

        /// <summary>
        /// Total thickness of all core layers (between first and last core layer indices, inclusive).
        /// </summary>
        private static double GetCoreWidth(CompoundStructure cs)
        {
            var layers = cs.GetLayers();
            int coreStart = cs.GetFirstCoreLayerIndex();
            int coreEnd = cs.GetLastCoreLayerIndex();
            double sum = 0.0;
            for (int i = coreStart; i <= coreEnd && i < layers.Count; i++)
            {
                sum += layers[i].Width;
            }
            return sum;
        }

        /// <summary>
        /// Returns true if the wall has a compound structure with more than one non-membrane layer.
        /// </summary>
        public static bool IsCompoundWall(Wall wall)
        {
            var cs = wall.WallType.GetCompoundStructure();
            if (cs == null) return false;

            int nonMembrane = cs.GetLayers().Count(l => l.Width > 1e-9);
            return nonMembrane > 1;
        }
    }
}
