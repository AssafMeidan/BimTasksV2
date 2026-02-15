using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Services;
using Microsoft.Win32;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Shows a SchedulePicker, exports the selected schedule to Excel using ScheduleExcelRoundtripService.
    /// </summary>
    public class ExportScheduleToExcelHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Get schedule: active view or picker
                ViewSchedule? schedule = GetSchedule(doc, uidoc);
                if (schedule == null) return;

                // Get save location
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    Title = "Export Schedule to Excel",
                    FileName = $"{schedule.Name}.xlsx"
                };

                if (saveDialog.ShowDialog() != true)
                    return;

                string filePath = saveDialog.FileName;

                // Export
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var service = container.Resolve<ScheduleExcelRoundtripService>();
                service.ExportScheduleToExcel(schedule, filePath);

                TaskDialog.Show("BimTasksV2",
                    $"Schedule '{schedule.Name}' exported successfully.\n\n{filePath}");

                Log.Information("ExportScheduleToExcel: Exported '{Name}' to {Path}",
                    schedule.Name, filePath);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ExportScheduleToExcel failed");
                TaskDialog.Show("BimTasksV2", $"Export failed:\n{ex.Message}");
            }
        }

        private ViewSchedule? GetSchedule(Document doc, UIDocument uidoc)
        {
            // If active view is a schedule, use it
            if (uidoc.ActiveView is ViewSchedule activeSchedule &&
                !activeSchedule.IsTitleblockRevisionSchedule &&
                !activeSchedule.IsInternalKeynoteSchedule)
            {
                return activeSchedule;
            }

            // Otherwise show picker
            var schedules = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTitleblockRevisionSchedule &&
                            !s.IsInternalKeynoteSchedule &&
                            !s.Definition.IsKeySchedule &&
                            !s.Name.StartsWith("<"))
                .ToList();

            if (schedules.Count == 0)
            {
                TaskDialog.Show("BimTasksV2", "The project contains no schedules.");
                return null;
            }

            var picker = new Views.SchedulePickerWindow(schedules);
            picker.ShowDialog();
            return picker.SelectedSchedule;
        }
    }
}
