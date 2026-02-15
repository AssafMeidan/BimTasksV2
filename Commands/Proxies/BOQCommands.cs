using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class BOQSetupCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "BOQSetupHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class BOQCreateSchedulesCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "BOQCreateSchedulesHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class BOQAutoFillQtyCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "BOQAutoFillQtyHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class BOQCalcEffUnitPriceCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "BOQCalcEffUnitPriceHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CreateSeifeiChozeCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CreateSeifeiChozeHandler"; }
}
