using Autodesk.Revit.UI;

namespace BimTasksV2.Services
{
    public interface IRevitContextService
    {
        UIApplication? UIApplication { get; set; }
        UIDocument? UIDocument { get; set; }
    }
}
