using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Services;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Exports a schedule to a temp Excel file, opens it in Excel for editing,
    /// waits for user confirmation, then imports changes back into the Revit model.
    /// </summary>
    public class EditScheduleInExcelHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // 1. Get schedule
                ViewSchedule? schedule = GetSchedule(doc, uidoc);
                if (schedule == null) return;

                // 2. Export to temp file
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var service = container.Resolve<ScheduleExcelRoundtripService>();

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string safeName = string.Join("_", schedule.Name.Split(Path.GetInvalidFileNameChars()));
                string filePath = Path.Combine(Path.GetTempPath(), $"BimTasksV2_{safeName}_{timestamp}.xlsx");

                service.ExportScheduleToExcel(schedule, filePath);

                // 3. Open in Excel
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });

                // 4. Wait for user confirmation
                var dlg = new TaskDialog("Edit Schedule in Excel")
                {
                    MainInstruction = "Schedule exported and opened in Excel.",
                    MainContent = "Edit the writable columns (white background) in Excel.\n" +
                                  "Read-only columns are grayed out.\n\n" +
                                  "When done, SAVE the file and CLOSE Excel,\n" +
                                  "then click 'Import Changes' to apply edits to the model.",
                    CommonButtons = TaskDialogCommonButtons.Cancel,
                    DefaultButton = TaskDialogResult.Cancel
                };
                dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Import Changes",
                    "Read the Excel file and apply changes to the Revit model");

                var result = dlg.Show();
                if (result != TaskDialogResult.CommandLink1)
                {
                    CleanupFile(filePath);
                    return;
                }

                // 5. Import changes
                Models.ImportResult importResult;
                using (var tx = new Transaction(doc, "Import Schedule Edits from Excel"))
                {
                    tx.Start();
                    try
                    {
                        importResult = service.ImportScheduleFromExcel(filePath, doc, schedule);
                        tx.Commit();
                    }
                    catch (IOException ioEx)
                    {
                        tx.RollBack();
                        TaskDialog.Show("File Locked",
                            "Cannot read the Excel file. Please close Excel and try again.\n\n" +
                            ioEx.Message);
                        return;
                    }
                    catch
                    {
                        tx.RollBack();
                        throw;
                    }
                }

                // 6. Show results
                ShowResults(importResult);

                // 7. Cleanup
                CleanupFile(filePath);

                Log.Information("EditScheduleInExcel: Completed for '{Name}'", schedule.Name);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "EditScheduleInExcel failed");
                TaskDialog.Show("BimTasksV2", $"Error:\n{ex.Message}");
            }
        }

        private ViewSchedule? GetSchedule(Document doc, UIDocument uidoc)
        {
            if (uidoc.ActiveView is ViewSchedule activeSchedule &&
                !activeSchedule.IsTitleblockRevisionSchedule &&
                !activeSchedule.IsInternalKeynoteSchedule)
            {
                return activeSchedule;
            }

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

        private void ShowResults(Models.ImportResult result)
        {
            string summary = $"Import Complete\n\n" +
                             $"Updated cells: {result.UpdatedCells}\n" +
                             $"Unchanged cells: {result.SkippedCells}\n" +
                             $"Failed cells: {result.FailedCells}\n" +
                             $"Elements affected: {result.AffectedElements}";

            if (result.Errors.Count > 0)
            {
                int showCount = Math.Min(result.Errors.Count, 10);
                summary += $"\n\nFirst {showCount} issues:";
                for (int i = 0; i < showCount; i++)
                    summary += $"\n  - {result.Errors[i]}";
                if (result.Errors.Count > showCount)
                    summary += $"\n  ... and {result.Errors.Count - showCount} more (see log file)";
            }

            TaskDialog.Show("Schedule Import Results", summary);
        }

        private void CleanupFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch { /* Best effort cleanup */ }
        }
    }
}
