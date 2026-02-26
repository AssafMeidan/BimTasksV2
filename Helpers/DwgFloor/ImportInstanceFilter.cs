using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace BimTasksV2.Helpers.DwgFloor
{
    public class ImportInstanceFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is ImportInstance;
        public bool AllowReference(Reference reference, XYZ position) => true;
    }
}
