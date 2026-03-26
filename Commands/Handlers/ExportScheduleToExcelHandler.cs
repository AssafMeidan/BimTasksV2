using System;
using System.Collections.Generic;
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
    /// Shows a multi-select SchedulePicker and exports selected schedules
    /// to one or multiple professionally formatted Excel files.
    /// </summary>
    public class ExportScheduleToExcelHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
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
                    return;
                }

                var picker = new Views.SchedulePickerWindow(
                    schedules,
                    title: "Export Schedule to Excel",
                    instruction: "Select schedules (Ctrl+Click for multiple):",
                    multiSelect: true);

                if (picker.ShowDialog() != true || picker.SelectedSchedules.Count == 0)
                    return;

                var selected = picker.SelectedSchedules;
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var service = container.Resolve<ScheduleExcelExportService>();

                if (picker.ExportAsSingleFile)
                {
                    ExportAsSingleFile(service, selected);
                }
                else
                {
                    ExportAsSeparateFiles(service, selected);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ExportScheduleToExcel failed");
                TaskDialog.Show("BimTasksV2", $"Export failed:\n{ex.Message}");
            }
        }

        private void ExportAsSingleFile(ScheduleExcelExportService service, List<ViewSchedule> selected)
        {
            string defaultName = selected.Count == 1
                ? $"{SafeName(selected[0].Name)}_{DateTime.Now:yyyyMMdd}.xlsx"
                : $"Schedules_{DateTime.Now:yyyyMMdd}.xlsx";

            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Save Schedules as Excel",
                FileName = defaultName
            };

            if (saveDialog.ShowDialog() != true)
                return;

            service.ExportSchedulesToExcel(selected, saveDialog.FileName);

            string names = string.Join(", ", selected.Select(s => s.Name));
            TaskDialog.Show("Export Complete",
                $"Exported {selected.Count} schedule(s) to:\n{saveDialog.FileName}");

            Log.Information("ExportScheduleToExcel: Exported {Count} schedules to {Path}",
                selected.Count, saveDialog.FileName);
        }

        private void ExportAsSeparateFiles(ScheduleExcelExportService service, List<ViewSchedule> selected)
        {
            // Ask for a folder
            using var folderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select folder for exported Excel files",
                UseDescriptionForTitle = true
            };

            if (folderDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                return;

            string folder = folderDialog.SelectedPath;
            int exported = 0;

            foreach (var schedule in selected)
            {
                string fileName = $"{SafeName(schedule.Name)}_{DateTime.Now:yyyyMMdd}.xlsx";
                string filePath = Path.Combine(folder, fileName);

                try
                {
                    service.ExportScheduleToExcel(schedule, filePath);
                    exported++;
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Failed to export schedule '{Name}'", schedule.Name);
                }
            }

            TaskDialog.Show("Export Complete",
                $"Exported {exported} of {selected.Count} schedule(s) to:\n{folder}");

            Log.Information("ExportScheduleToExcel: Exported {Exported}/{Total} schedules to {Folder}",
                exported, selected.Count, folder);
        }

        private static string SafeName(string name)
        {
            return string.Join("_",
                name.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));
        }
    }
}
