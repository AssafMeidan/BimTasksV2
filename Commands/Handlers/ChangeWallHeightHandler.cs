using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Shows the ChangeWallHeightDialog and applies height changes to selected walls.
    /// The dialog is created via the DI container to support ViewModelLocator auto-wiring.
    /// </summary>
    public class ChangeWallHeightHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Update context service so the VM can access UIApplication
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var contextService = container.Resolve<Services.IRevitContextService>();
                contextService.UIApplication = uiApp;
                contextService.UIDocument = uidoc;

                // Create the dialog window via container for VM auto-wiring
                var dialog = container.Resolve<Views.ChangeWallHeightDialog>();
                dialog.ShowDialog();

                Log.Information("ChangeWallHeight: Dialog closed");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ChangeWallHeight failed");
                TaskDialog.Show("BimTasksV2", $"Error: {ex.Message}");
            }
        }
    }
}
