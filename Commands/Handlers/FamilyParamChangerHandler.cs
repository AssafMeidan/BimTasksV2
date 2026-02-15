using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Stub handler - OldApp version was incomplete.
    /// </summary>
    public class FamilyParamChangerHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            TaskDialog.Show("Family Parameter Changer", "This feature is not yet implemented.");
        }
    }
}
