using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Shows the CopyCategoryFromLinkView dialog for copying categories from linked models.
    /// </summary>
    public class CopyCategoryFromLinkHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;

                // Update context service
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uiApp.ActiveUIDocument;

                // Resolve and show the dialog
                var dialog = container.Resolve<Views.CopyCategoryFromLinkView>();
                dialog.ShowDialog();

                Log.Information("CopyCategoryFromLink: Dialog closed");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CopyCategoryFromLink failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
