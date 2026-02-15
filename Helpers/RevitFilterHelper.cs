using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Static helper for common Revit element filtering operations.
    /// </summary>
    public static class RevitFilterHelper
    {
        /// <summary>
        /// Returns all elements of the requested type from the document.
        /// </summary>
        /// <typeparam name="T">The element type to collect (must derive from Element).</typeparam>
        /// <param name="doc">The Revit document to query.</param>
        /// <returns>A list of elements matching the requested type.</returns>
        public static IList<T> GetElementsOfType<T>(Document doc) where T : Element
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(T))
                .Cast<T>()
                .ToList();
        }

        /// <summary>
        /// Returns all elements of the requested System.Type matching the given built-in category.
        /// </summary>
        /// <param name="doc">The Revit document to query.</param>
        /// <param name="type">The System.Type to filter by.</param>
        /// <param name="bic">The built-in category to filter by.</param>
        /// <returns>A FilteredElementCollector with the applied filters.</returns>
        public static FilteredElementCollector GetElementsOfType(Document doc, Type type, BuiltInCategory bic)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(bic)
                .OfClass(type);
        }

        /// <summary>
        /// Returns all element instances (not types) of the given built-in category.
        /// </summary>
        /// <param name="doc">The Revit document to query.</param>
        /// <param name="bic">The built-in category to filter by.</param>
        /// <returns>A list of elements matching the category.</returns>
        public static IList<Element> GetInstancesOfCategory(Document doc, BuiltInCategory bic)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
                .ToElements();
        }
    }
}
