using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class TestInfrastructureCommand : Infrastructure.ProxyCommand
    {
        protected override string HandlerTypeName => "TestInfrastructureRunner";
    }
}
