using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class ExportScheduleToExcelCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ExportScheduleToExcelHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class EditScheduleInExcelCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "EditScheduleInExcelHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ImportKeySchedulesFromExcelCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ImportKeySchedulesFromExcelHandler"; }
}
