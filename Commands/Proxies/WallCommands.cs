using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class SetSelectedWindowAndDoorPhaseToWallPhaseCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "SetSelectedWindowAndDoorPhaseToWallPhaseHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CopyLinkedWallsCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CopyLinkedWallsHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class AdjustWallHeightToFloorsCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "AdjustWallHeightToFloorsHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ChangeWallHeightCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ChangeWallHeightHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class AddChipuyToWallCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "AddChipuyToWallHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class AddChipuyToExternalWallCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "AddChipuyToExternalWallHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class AddChipuyToInternalWallCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "AddChipuyToInternalWallHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CreateWindowFamiliesCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CreateWindowFamiliesHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class SplitWallCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "SplitWallHandler"; }

}
