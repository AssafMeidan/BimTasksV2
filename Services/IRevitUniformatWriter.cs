using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Scope for applying Uniformat codes: Types or Instances.
    /// </summary>
    public enum ApplyScope
    {
        Types,
        Instances
    }

    /// <summary>
    /// Target mode: apply to current Selection or by Categories.
    /// </summary>
    public enum TargetMode
    {
        Selection,
        Categories
    }

    /// <summary>
    /// Writes Uniformat assembly codes and names into shared parameters on Revit elements.
    /// </summary>
    public interface IRevitUniformatWriter
    {
        /// <summary>
        /// Ensures shared parameters and bindings exist for the target categories.
        /// </summary>
        void EnsureParameters(Document doc, IEnumerable<BuiltInCategory> targets, bool bindToTypes);

        /// <summary>
        /// Applies Uniformat code and name to elements. Returns the count of updated elements.
        /// </summary>
        int Apply(
            UIDocument uidoc,
            string uniformatCode,
            string uniformatName,
            ApplyScope scope,
            TargetMode targetMode,
            IEnumerable<BuiltInCategory> categories);
    }
}
