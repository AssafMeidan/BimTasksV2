using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.FloorSplitter
{
    /// <summary>
    /// Data describing a single compound floor layer with its vertical offset.
    /// </summary>
    public class FloorLayerInfo
    {
        public int Index { get; set; }
        public double Thickness { get; set; }
        public ElementId MaterialId { get; set; }
        public string MaterialName { get; set; } = "";
        public MaterialFunctionAssignment Function { get; set; }
        public bool IsStructural => Function == MaterialFunctionAssignment.Structure;

        /// <summary>
        /// Vertical offset from the original floor's top surface to this layer's top.
        /// Always negative or zero (layers stack downward from the top).
        /// </summary>
        public double TopOffset { get; set; }

        /// <summary>
        /// The FloorType to use for the replacement floor (resolved or auto-created).
        /// </summary>
        public FloorType? ResolvedType { get; set; }

        public bool IsAutoCreatedType { get; set; }
    }

    /// <summary>
    /// Reads CompoundStructure layers from a floor type and computes
    /// each layer's vertical offset from the floor's top surface.
    /// Layers are ordered top→bottom (index 0 = topmost).
    /// </summary>
    public static class FloorLayerAnalyzer
    {
        /// <summary>
        /// Analyzes a compound floor and returns info for each non-membrane layer.
        /// </summary>
        public static List<FloorLayerInfo> AnalyzeLayers(Floor floor)
        {
            var result = new List<FloorLayerInfo>();
            var doc = floor.Document;
            var floorType = doc.GetElement(floor.GetTypeId()) as FloorType;
            if (floorType == null) return result;

            var cs = floorType.GetCompoundStructure();
            if (cs == null) return result;

            var layers = cs.GetLayers();
            if (layers.Count <= 1) return result;

            // Running distance from the top surface downward
            double runningOffset = 0.0;

            for (int i = 0; i < layers.Count; i++)
            {
                var layer = layers[i];
                double thickness = layer.Width;

                // Skip zero-thickness membrane layers
                if (thickness < 1e-9)
                    continue;

                string materialName = "";
                if (layer.MaterialId != ElementId.InvalidElementId)
                {
                    var mat = doc.GetElement(layer.MaterialId) as Material;
                    if (mat != null)
                        materialName = mat.Name;
                }

                result.Add(new FloorLayerInfo
                {
                    Index = i,
                    Thickness = thickness,
                    MaterialId = layer.MaterialId,
                    MaterialName = materialName,
                    Function = layer.Function,
                    TopOffset = -runningOffset // Negative = downward from top
                });

                runningOffset += thickness;
            }

            return result;
        }

        /// <summary>
        /// Returns true if the floor has a compound structure with more than one non-membrane layer.
        /// </summary>
        public static bool IsCompoundFloor(Floor floor)
        {
            var doc = floor.Document;
            var floorType = doc.GetElement(floor.GetTypeId()) as FloorType;
            if (floorType == null) return false;

            var cs = floorType.GetCompoundStructure();
            if (cs == null) return false;

            int nonMembrane = cs.GetLayers().Count(l => l.Width > 1e-9);
            return nonMembrane > 1;
        }
    }
}
