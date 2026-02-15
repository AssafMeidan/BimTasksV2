using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Selection filter that only allows the user to pick RevitLinkInstance elements.
    /// Used when prompting the user to select a linked Revit model.
    /// </summary>
    public class LinkSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is RevitLinkInstance;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
