using System;

namespace BimTasksV2.Services
{
    public interface ICommandDispatcherService
    {
        void Enqueue(Action<Autodesk.Revit.UI.UIApplication> action);
    }
}
