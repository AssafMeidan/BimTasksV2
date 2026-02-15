using Autodesk.Revit.UI;

namespace BimTasksV2.Commands.Infrastructure
{
    public interface ICommandHandler
    {
        void Execute(UIApplication uiApp);
    }
}
