using Autodesk.Revit.Attributes;

namespace BimTasksV2.Commands.Proxies
{
    [Transaction(TransactionMode.Manual)]
    public class ShowFilterTreeWindowCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ShowFilterTreeWindowHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CheckSelectedWallsAreaAndVolumeCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CheckSelectedWallsAreaAndVolumeHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ShowUniformatWindowCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ShowUniformatWindowHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class CopyCategoryFromLinkCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "CopyCategoryFromLinkHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ToggleDockablePanelCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ToggleDockablePanelHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ToggleToolbarCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ToggleToolbarHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class WebAppDataExtractorCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "WebAppDataExtractorHandler"; }

    [Transaction(TransactionMode.Manual)]
    public class ToggleVoskVoiceRecognitionCommand : Infrastructure.ProxyCommand
    { protected override string HandlerTypeName => "ToggleVoskVoiceRecognitionHandler"; }
}
