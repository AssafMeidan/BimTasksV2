using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class SetBeamsToGroundZeroCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "SetBeamsToGroundZeroHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class JoinAllConcreteWallsAndFloorsGeometryCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "JoinAllConcreteWallsAndFloorsGeometryHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class SplitFloorCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "SplitFloorHandler"; }
}
