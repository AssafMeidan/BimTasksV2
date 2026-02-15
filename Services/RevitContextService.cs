using Autodesk.Revit.UI;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Singleton service that holds a reference to the current Revit UIApplication.
    /// UIDocument is computed from UIApplication.ActiveUIDocument so it always
    /// reflects the currently-active document (critical bug fix over OldApp which
    /// stored a stale UIDocument reference).
    /// </summary>
    public sealed class RevitContextService : IRevitContextService
    {
        public UIApplication? UIApplication { get; set; }

        public UIDocument? UIDocument
        {
            get => UIApplication?.ActiveUIDocument;
            set { /* computed from UIApplication.ActiveUIDocument — setter kept for interface compat */ }
        }
    }
}
