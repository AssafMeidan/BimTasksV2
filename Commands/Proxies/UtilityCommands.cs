using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class JoinSelectedBeamsToBlockWallsCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "JoinSelectedBeamsToBlockWallsHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class AddClunasFromDwgCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "AddClunasFromDwgHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CreateSelectedItemCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CreateSelectedItemHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class FamilyParamChangerCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "FamilyParamChangerHandler"; }
}
