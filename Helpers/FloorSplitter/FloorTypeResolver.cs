using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.FloorSplitter
{
    /// <summary>
    /// Finds or auto-creates single-layer FloorTypes that match a given layer's
    /// thickness and material, for use when splitting compound floors.
    /// </summary>
    public static class FloorTypeResolver
    {
        private const double Tolerance = 1e-9;

        /// <summary>
        /// Finds an existing single-layer FloorType or auto-creates one matching the layer.
        /// </summary>
        public static FloorType FindOrCreateType(Document doc, FloorLayerInfo layer)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (layer == null) throw new ArgumentNullException(nameof(layer));

            var existingMatch = FindExactMatch(doc, layer.Thickness, layer.MaterialId);
            if (existingMatch != null)
            {
                layer.IsAutoCreatedType = false;
                layer.ResolvedType = existingMatch;
                Log.Information("[FloorTypeResolver] Found existing type '{TypeName}' for layer {Index}",
                    existingMatch.Name, layer.Index);
                return existingMatch;
            }

            var newType = CreateSingleLayerType(doc, layer);
            layer.IsAutoCreatedType = true;
            layer.ResolvedType = newType;
            Log.Information("[FloorTypeResolver] Auto-created type '{TypeName}' for layer {Index}",
                newType.Name, layer.Index);
            return newType;
        }

        /// <summary>
        /// Returns all FloorTypes whose total width matches the given thickness.
        /// </summary>
        public static List<FloorType> FindMatchingTypes(Document doc, double thickness)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            return new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .Where(ft =>
                {
                    var cs = ft.GetCompoundStructure();
                    if (cs == null) return false;
                    double totalWidth = cs.GetLayers().Sum(l => l.Width);
                    return Math.Abs(totalWidth - thickness) < Tolerance;
                })
                .OrderBy(ft => ft.Name)
                .ToList();
        }

        private static FloorType? FindExactMatch(Document doc, double thickness, ElementId materialId)
        {
            var floorTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>();

            foreach (var ft in floorTypes)
            {
                var cs = ft.GetCompoundStructure();
                if (cs == null) continue;

                var layers = cs.GetLayers();
                if (layers.Count != 1) continue;

                var singleLayer = layers[0];
                if (Math.Abs(singleLayer.Width - thickness) > Tolerance) continue;
                if (singleLayer.MaterialId != materialId) continue;

                return ft;
            }

            return null;
        }

        private static FloorType CreateSingleLayerType(Document doc, FloorLayerInfo layer)
        {
            var templateType = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .FirstOrDefault(ft => ft.GetCompoundStructure() != null);

            if (templateType == null)
                throw new InvalidOperationException("No floor type with compound structure found to use as template.");

            double thicknessMm = Math.Round(
                UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters), 0);
            string safeMaterialName = SanitizeName(layer.MaterialName);
            string newName = $"Split_{safeMaterialName}_{thicknessMm:F0}mm";
            newName = EnsureUniqueName(doc, newName);

            var newType = templateType.Duplicate(newName) as FloorType;
            if (newType == null)
                throw new InvalidOperationException($"Failed to duplicate floor type '{templateType.Name}' as '{newName}'.");

            var newCs = newType.GetCompoundStructure();
            if (newCs == null)
            {
                newCs = CompoundStructure.CreateSingleLayerCompoundStructure(
                    MaterialFunctionAssignment.Structure, layer.Thickness, layer.MaterialId);
            }
            else
            {
                var newLayers = new List<CompoundStructureLayer>
                {
                    new CompoundStructureLayer(layer.Thickness, layer.Function, layer.MaterialId)
                };
                newCs.SetLayers(newLayers);
                newCs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                newCs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);
            }

            newType.SetCompoundStructure(newCs);

            Log.Information("[FloorTypeResolver] Created new FloorType '{TypeName}' (Id={TypeId})",
                newName, newType.Id);
            return newType;
        }

        private static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "NoMaterial";

            var invalid = new[] { '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' };
            foreach (char c in invalid)
                name = name.Replace(c, '_');

            return name.Trim();
        }

        private static string EnsureUniqueName(Document doc, string baseName)
        {
            var existingNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(FloorType))
                    .Cast<FloorType>()
                    .Select(ft => ft.Name),
                StringComparer.OrdinalIgnoreCase);

            if (!existingNames.Contains(baseName))
                return baseName;

            int suffix = 2;
            while (existingNames.Contains($"{baseName}_{suffix}"))
                suffix++;

            string uniqueName = $"{baseName}_{suffix}";
            Log.Warning("[FloorTypeResolver] Name '{BaseName}' already exists; using '{UniqueName}'",
                baseName, uniqueName);
            return uniqueName;
        }
    }
}
