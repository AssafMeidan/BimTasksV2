using Autodesk.Revit.DB;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers.WallSplitter
{
    /// <summary>
    /// Finds or auto-creates single-layer WallTypes that match a given layer's
    /// thickness and material, for use when splitting compound walls into individual layers.
    /// </summary>
    public static class WallTypeResolver
    {
        private const double Tolerance = 1e-9;

        /// <summary>
        /// Searches for an existing single-layer WallType whose layer width and material
        /// match the given <paramref name="layer"/>. If none is found, duplicates a generic
        /// basic wall type, renames it, and sets up a single-layer CompoundStructure.
        /// </summary>
        /// <param name="doc">The Revit document (must be in an active transaction).</param>
        /// <param name="layer">The layer info describing the desired thickness and material.</param>
        /// <returns>The matching or newly created WallType.</returns>
        public static WallType FindOrCreateType(Document doc, LayerInfo layer)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (layer == null) throw new ArgumentNullException(nameof(layer));

            // Try to find an existing single-layer WallType that matches
            var existingMatch = FindExactMatch(doc, layer.Thickness, layer.MaterialId);
            if (existingMatch != null)
            {
                layer.IsAutoCreatedType = false;
                layer.ResolvedType = existingMatch;
                Log.Information("[WallTypeResolver] Found existing type '{TypeName}' for layer {Index} ({Material}, {Thickness:F1}mm)",
                    existingMatch.Name, layer.Index, layer.MaterialName,
                    UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters));
                return existingMatch;
            }

            // No match found — auto-create by duplicating a generic basic wall type
            var newType = CreateSingleLayerType(doc, layer);
            layer.IsAutoCreatedType = true;
            layer.ResolvedType = newType;
            Log.Information("[WallTypeResolver] Auto-created type '{TypeName}' for layer {Index} ({Material}, {Thickness:F1}mm)",
                newType.Name, layer.Index, layer.MaterialName,
                UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters));
            return newType;
        }

        /// <summary>
        /// Returns all WallTypes whose total width matches the given thickness (within tolerance).
        /// Useful for populating dialog dropdowns where the user can pick an alternative type.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="thickness">Target thickness in internal units (feet).</param>
        /// <returns>List of matching WallTypes.</returns>
        public static List<WallType> FindMatchingTypes(Document doc, double thickness)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            return new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .Where(wt => wt.Kind == WallKind.Basic)
                .Where(wt => Math.Abs(wt.Width - thickness) < Tolerance)
                .OrderBy(wt => wt.Name)
                .ToList();
        }

        /// <summary>
        /// Searches for an existing basic WallType with exactly one compound layer
        /// whose width and material match the specified values.
        /// </summary>
        private static WallType? FindExactMatch(Document doc, double thickness, ElementId materialId)
        {
            var wallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .Where(wt => wt.Kind == WallKind.Basic);

            foreach (var wt in wallTypes)
            {
                var cs = wt.GetCompoundStructure();
                if (cs == null)
                    continue;

                var layers = cs.GetLayers();
                if (layers.Count != 1)
                    continue;

                var singleLayer = layers[0];
                if (Math.Abs(singleLayer.Width - thickness) > Tolerance)
                    continue;

                if (singleLayer.MaterialId != materialId)
                    continue;

                return wt;
            }

            return null;
        }

        /// <summary>
        /// Creates a new single-layer WallType by duplicating the first available basic wall type,
        /// renaming it with the convention "Split_{MaterialName}_{thickness_in_mm}mm",
        /// and setting up a single-layer CompoundStructure.
        /// </summary>
        private static WallType CreateSingleLayerType(Document doc, LayerInfo layer)
        {
            // Find a basic wall type to use as a template for duplication
            var templateType = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .FirstOrDefault(wt => wt.Kind == WallKind.Basic);

            if (templateType == null)
                throw new InvalidOperationException("No basic wall type found in the document to use as a template for duplication.");

            // Build the new type name
            double thicknessMm = Math.Round(
                UnitUtils.ConvertFromInternalUnits(layer.Thickness, UnitTypeId.Millimeters), 0);
            string safeMaterialName = SanitizeName(layer.MaterialName);
            string newName = $"Split_{safeMaterialName}_{thicknessMm:F0}mm";

            // Ensure the name is unique — append a suffix if needed
            newName = EnsureUniqueName(doc, newName);

            // Duplicate the template type
            var newType = templateType.Duplicate(newName) as WallType;
            if (newType == null)
                throw new InvalidOperationException($"Failed to duplicate wall type '{templateType.Name}' as '{newName}'.");

            // Build a single-layer CompoundStructure
            var newCs = newType.GetCompoundStructure();
            if (newCs == null)
            {
                // Unlikely for a basic wall, but handle defensively
                newCs = CompoundStructure.CreateSingleLayerCompoundStructure(
                    MaterialFunctionAssignment.Structure, layer.Thickness, layer.MaterialId);
            }
            else
            {
                // Clear existing layers and set up a single layer
                var newLayers = new List<CompoundStructureLayer>
                {
                    new CompoundStructureLayer(layer.Thickness, layer.Function, layer.MaterialId)
                };
                newCs.SetLayers(newLayers);
                newCs.SetNumberOfShellLayers(ShellLayerType.Exterior, 0);
                newCs.SetNumberOfShellLayers(ShellLayerType.Interior, 0);
            }

            newType.SetCompoundStructure(newCs);

            Log.Information("[WallTypeResolver] Created new WallType '{TypeName}' (Id={TypeId})",
                newName, newType.Id);

            return newType;
        }

        /// <summary>
        /// Removes characters that are invalid in Revit type names.
        /// </summary>
        private static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "NoMaterial";

            // Remove characters that Revit disallows in type names: { } [ ] | ; < > ? ` ~
            var invalid = new[] { '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' };
            foreach (char c in invalid)
            {
                name = name.Replace(c, '_');
            }

            return name.Trim();
        }

        /// <summary>
        /// Ensures the given name is unique among existing WallType names in the document.
        /// Appends _2, _3, etc. if a collision is found.
        /// </summary>
        private static string EnsureUniqueName(Document doc, string baseName)
        {
            var existingNames = new HashSet<string>(
                new FilteredElementCollector(doc)
                    .OfClass(typeof(WallType))
                    .Cast<WallType>()
                    .Select(wt => wt.Name),
                StringComparer.OrdinalIgnoreCase);

            if (!existingNames.Contains(baseName))
                return baseName;

            int suffix = 2;
            while (existingNames.Contains($"{baseName}_{suffix}"))
            {
                suffix++;
            }

            string uniqueName = $"{baseName}_{suffix}";
            Log.Warning("[WallTypeResolver] Name '{BaseName}' already exists; using '{UniqueName}' instead",
                baseName, uniqueName);
            return uniqueName;
        }
    }
}
