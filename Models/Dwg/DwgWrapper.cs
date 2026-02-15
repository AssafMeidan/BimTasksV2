using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Models.Dwg
{
    /// <summary>
    /// Wraps a DWG ImportInstance and extracts arc geometry from it.
    /// Used for placing piles/columns at arc center points found on
    /// a specific DWG layer (e.g., "fnd" for foundation piles).
    /// </summary>
    public class DwgWrapper
    {
        private readonly Document _doc;
        private readonly GeometryElement _geoElement;

        /// <summary>
        /// The selected DWG import instance.
        /// </summary>
        public ImportInstance SelectedDwg { get; }

        /// <summary>
        /// Extracted arc geometry objects from the DWG, filtered by layer
        /// and de-duplicated by center point (keeping the largest radius).
        /// </summary>
        public List<Arc> GeometryArcs { get; } = new List<Arc>();

        /// <summary>
        /// The DWG layer name to filter arcs from.
        /// </summary>
        public string LayerName { get; set; } = "fnd";

        /// <summary>
        /// Creates a DwgWrapper for the given ImportInstance.
        /// Automatically extracts arc geometry from the default layer.
        /// </summary>
        /// <param name="selectedDwg">The DWG ImportInstance element.</param>
        public DwgWrapper(ImportInstance selectedDwg)
        {
            _doc = selectedDwg.Document;
            SelectedDwg = selectedDwg;

            var options = new Options { IncludeNonVisibleObjects = false };
            _geoElement = selectedDwg.get_Geometry(options);

            ExtractArcsFromGeometry();
        }

        /// <summary>
        /// Creates a DwgWrapper with a custom layer name filter.
        /// </summary>
        /// <param name="selectedDwg">The DWG ImportInstance element.</param>
        /// <param name="layerName">The DWG layer name to extract arcs from.</param>
        public DwgWrapper(ImportInstance selectedDwg, string layerName)
            : this(selectedDwg)
        {
            LayerName = layerName;
            GeometryArcs.Clear();
            ExtractArcsFromGeometry();
        }

        /// <summary>
        /// Extracts arc geometry objects from the DWG instance geometry,
        /// filtering by the specified layer name. When multiple arcs share
        /// the same center point, only the one with the largest radius is kept.
        /// </summary>
        private void ExtractArcsFromGeometry()
        {
            if (_geoElement == null)
                return;

            foreach (var geometryObject in _geoElement)
            {
                if (geometryObject is GeometryInstance geometryInstance)
                {
                    var instanceGeometry = geometryInstance.GetInstanceGeometry();
                    if (instanceGeometry == null)
                        continue;

                    foreach (var obj in instanceGeometry)
                    {
                        if (obj is Arc arc)
                        {
                            // Filter by DWG layer name via GraphicsStyle
                            var graphicsStyle = _doc.GetElement(obj.GraphicsStyleId) as GraphicsStyle;
                            if (graphicsStyle?.GraphicsStyleCategory?.Name != LayerName)
                                continue;

                            // De-duplicate by center point: keep the arc with the largest radius
                            bool shouldAdd = true;
                            for (int i = GeometryArcs.Count - 1; i >= 0; i--)
                            {
                                var existing = GeometryArcs[i];
                                if (arc.Center.X == existing.Center.X && arc.Center.Y == existing.Center.Y)
                                {
                                    if (arc.Radius > existing.Radius)
                                    {
                                        GeometryArcs.RemoveAt(i);
                                    }
                                    else
                                    {
                                        shouldAdd = false;
                                    }
                                    break;
                                }
                            }

                            if (shouldAdd)
                            {
                                GeometryArcs.Add(arc);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Returns the center points of all extracted arcs.
        /// </summary>
        public IList<XYZ> GetArcCenterPoints()
        {
            return GeometryArcs.Select(a => a.Center).ToList();
        }

        /// <summary>
        /// Returns a list of (Center, Radius) tuples for all extracted arcs.
        /// </summary>
        public IList<(XYZ Center, double Radius)> GetArcData()
        {
            return GeometryArcs.Select(a => (a.Center, a.Radius)).ToList();
        }
    }
}
