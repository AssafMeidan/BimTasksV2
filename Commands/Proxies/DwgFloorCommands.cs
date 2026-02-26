using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class PickAndCreateFloorCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "PickAndCreateFloorHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ImportAllDwgFloorsCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ImportAllDwgFloorsHandler"; }
}
