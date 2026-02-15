using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using Microsoft.Win32;
using OfficeOpenXml;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Creates Seifei Choze (contract item / BOQ summary) key schedules from Excel data.
    /// Imports key schedule data and creates/updates key schedules for target categories.
    /// </summary>
    public class CreateSeifeiChozeHandler : ICommandHandler
    {
        #region Configuration

        private static readonly BuiltInCategory[] TargetCategories = new[]
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Ceilings,
        };

        private static readonly List<string> RequiredParamNames = new List<string>
        {
            "\u05E1\u05E2\u05D9\u05E3 \u05EA\u05E7\u05E6\u05D9\u05D1\u05D9",  // "סעיף תקציבי"
            "\u05EA\u05D0\u05D5\u05E8 \u05E1\u05E2\u05D9\u05E3",                // "תאור סעיף"
            "\u05E7\u05D1\u05DC\u05DF \u05DE\u05E9\u05E0\u05D4",                // "קבלן משנה"
            "\u05DE\u05D7\u05D9\u05E8 \u05DB\u05D5\u05DC\u05DC"                 // "מחיר כולל"
        };

        private const int MaxExcelRows = 10000;
        private const int MaxConsecutiveEmptyRows = 5;

        #endregion

        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            Log.Information("=== Starting CreateSeifeiChoze ===");

            try
            {
                // Step 1: Ensure shared parameters are bound
                if (!EnsureProjectParameters(doc, out string? paramError))
                {
                    Log.Warning("Parameter binding failed: {Error}", paramError);
                    TaskDialog.Show("BimTasksV2", paramError ?? "Parameter binding failed.");
                    return;
                }

                // Step 2: Get Excel file
                string? excelPath = GetExcelFilePath();
                if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                {
                    Log.Information("User cancelled file selection or file not found");
                    return;
                }

                // Step 3: Read Excel data
                var (excelData, readErrors) = ReadExcelDataSafe(excelPath!);
                if (excelData.Count == 0)
                {
                    string errorMsg = "The selected Excel file contains no valid data.";
                    if (readErrors.Any())
                        errorMsg += "\n\nErrors:\n" + string.Join("\n", readErrors);
                    TaskDialog.Show("BimTasksV2 - Error", errorMsg);
                    return;
                }

                Log.Information("Read {Count} entries from Excel", excelData.Count);

                // Step 4: Create/Update key schedules
                var (successCount, failures) = UpdateAllKeySchedules(doc, excelData);

                // Step 5: Report results
                ReportResults(excelData.Count, successCount, failures, readErrors);

                Log.Information("=== CreateSeifeiChoze completed. Success: {Success}/{Total} ===",
                    successCount, TargetCategories.Length);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "CreateSeifeiChoze failed");
                TaskDialog.Show("BimTasksV2", $"An unexpected error occurred:\n\n{ex.Message}");
            }
        }

        #region Parameter Binding

        private bool EnsureProjectParameters(Document doc, out string? error)
        {
            error = null;

            if (AreAllParametersBound(doc))
            {
                Log.Debug("All parameters already bound");
                return true;
            }

            var app = doc.Application;
            var sharedFile = app.OpenSharedParameterFile();
            if (sharedFile == null)
            {
                error = "No Shared Parameter File is linked to Revit.\n\n" +
                        "Please go to: Manage > Shared Parameters\n" +
                        "And browse to your BIMTasks shared parameter .txt file.";
                return false;
            }

            var catSet = app.Create.NewCategorySet();
            foreach (var bic in TargetCategories)
            {
                var cat = doc.Settings.Categories.get_Item(bic);
                if (cat != null && cat.AllowsBoundParameters)
                    catSet.Insert(cat);
            }

            if (catSet.Size == 0)
            {
                error = "No valid categories found for parameter binding.";
                return false;
            }

            using (var tx = new Transaction(doc, "Bind BIMTasks Parameters"))
            {
                tx.Start();

                foreach (string paramName in RequiredParamNames)
                {
                    var definition = FindDefinitionInSharedFile(sharedFile, paramName);
                    if (definition == null)
                    {
                        error = $"Parameter '{paramName}' was NOT found in your Shared Parameter file.";
                        tx.RollBack();
                        return false;
                    }

                    var binding = app.Create.NewInstanceBinding(catSet);
                    try
                    {
                        if (!doc.ParameterBindings.Contains(definition))
                            doc.ParameterBindings.Insert(definition, binding);
                        else
                            doc.ParameterBindings.ReInsert(definition, binding);
                    }
                    catch (Exception ex)
                    {
                        error = $"Failed to bind parameter '{paramName}':\n{ex.Message}";
                        tx.RollBack();
                        return false;
                    }
                }

                tx.Commit();
            }

            return true;
        }

        private bool AreAllParametersBound(Document doc)
        {
            if (!IsParameterBound(doc, RequiredParamNames[0], TargetCategories[0]))
                return false;
            if (!IsParameterBound(doc, RequiredParamNames[^1], TargetCategories[^1]))
                return false;
            return true;
        }

        private bool IsParameterBound(Document doc, string paramName, BuiltInCategory bic)
        {
            var targetCat = doc.Settings.Categories.get_Item(bic);
            if (targetCat == null) return false;

            var map = doc.ParameterBindings;
            var it = map.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                if (it.Key.Name == paramName)
                {
                    if (it.Current is ElementBinding binding && binding.Categories.Contains(targetCat))
                        return true;
                }
            }
            return false;
        }

        private Definition? FindDefinitionInSharedFile(DefinitionFile sharedFile, string name)
        {
            foreach (DefinitionGroup group in sharedFile.Groups)
            {
                foreach (Definition def in group.Definitions)
                {
                    if (def.Name == name) return def;
                }
            }
            return null;
        }

        #endregion

        #region Key Schedule Operations

        private (int successCount, List<string> failures) UpdateAllKeySchedules(Document doc, List<KeyEntry> data)
        {
            var failures = new List<string>();
            int successCount = 0;

            using (var tx = new Transaction(doc, "Import Key Schedules"))
            {
                tx.Start();

                foreach (var bic in TargetCategories)
                {
                    string catName = LabelUtils.GetLabelFor(bic);
                    try
                    {
                        var schedule = GetOrCreateKeySchedule(doc, bic);
                        if (schedule == null)
                        {
                            failures.Add($"{catName}: Cannot create key schedule");
                            continue;
                        }

                        int updated = UpdateScheduleRows(doc, schedule, data);
                        Log.Information("Updated {Count} rows in '{Name}'", updated, schedule.Name);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failures.Add($"{catName}: {ex.Message}");
                        Log.Warning(ex, "Failed for category {Category}", catName);
                    }
                }

                if (successCount > 0)
                    tx.Commit();
                else
                    tx.RollBack();
            }

            return (successCount, failures);
        }

        private ViewSchedule? GetOrCreateKeySchedule(Document doc, BuiltInCategory bic)
        {
            var cat = Category.GetCategory(doc, bic);
            if (cat == null) return null;

            string catName = LabelUtils.GetLabelFor(bic);
            string scheduleName = $"BIMTasks_Key_{catName}";

            var schedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(v => v.Name.Equals(scheduleName, StringComparison.OrdinalIgnoreCase)
                                     && v.Definition.IsKeySchedule);

            if (schedule == null)
            {
                try
                {
                    schedule = ViewSchedule.CreateKeySchedule(doc, cat.Id);
                    schedule.Name = scheduleName;
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    return null;
                }
            }

            EnsureScheduleFields(doc, schedule);
            return schedule;
        }

        private void EnsureScheduleFields(Document doc, ViewSchedule schedule)
        {
            var definition = schedule.Definition;
            var existingFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < definition.GetFieldCount(); i++)
                existingFields.Add(definition.GetField(i).GetName());

            foreach (var paramName in RequiredParamNames)
            {
                if (existingFields.Contains(paramName)) continue;

                var field = definition.GetSchedulableFields()
                    .FirstOrDefault(sf => sf.GetName(doc) == paramName);

                if (field != null)
                    definition.AddField(field);
            }
        }

        private int UpdateScheduleRows(Document doc, ViewSchedule schedule, List<KeyEntry> data)
        {
            var tableData = schedule.GetTableData();
            var sectionData = tableData.GetSectionData(SectionType.Body);

            var existingKeys = new FilteredElementCollector(doc, schedule.Id)
                .ToElements()
                .Where(e => !string.IsNullOrEmpty(e.Name))
                .GroupBy(e => e.Name)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            int updatedCount = 0;

            foreach (var entry in data)
            {
                Element? keyElement;

                if (!existingKeys.TryGetValue(entry.KeyName, out keyElement))
                {
                    try
                    {
                        sectionData.InsertRow(sectionData.FirstRowNumber);
                        var allElements = new FilteredElementCollector(doc, schedule.Id).ToElements();
                        keyElement = allElements.FirstOrDefault(e =>
                            string.IsNullOrEmpty(e.Name) ||
                            (!existingKeys.ContainsKey(e.Name) && e.Name == ""));

                        if (keyElement != null)
                            existingKeys[entry.KeyName] = keyElement;
                    }
                    catch (Exception ex)
                    {
                        Log.Warning(ex, "Failed to insert row for key '{KeyName}'", entry.KeyName);
                        continue;
                    }
                }

                if (keyElement != null)
                {
                    SetParamSafe(keyElement, "Key Name", entry.KeyName);
                    SetParamSafe(keyElement, RequiredParamNames[0], entry.ItemNumber);
                    SetParamSafe(keyElement, RequiredParamNames[1], entry.Description);
                    SetParamSafe(keyElement, RequiredParamNames[2], entry.Subcontractor);
                    SetParamSafe(keyElement, RequiredParamNames[3], entry.Price);
                    updatedCount++;
                }
            }

            return updatedCount;
        }

        #endregion

        #region Excel Reading

        private (List<KeyEntry> data, List<string> errors) ReadExcelDataSafe(string path)
        {
            var list = new List<KeyEntry>();
            var errors = new List<string>();

            try
            {
                byte[] fileBytes;
                try
                {
                    fileBytes = File.ReadAllBytes(path);
                }
                catch (IOException ioEx)
                {
                    errors.Add($"Cannot read file (may be open in Excel): {ioEx.Message}");
                    return (list, errors);
                }

                using (var stream = new MemoryStream(fileBytes))
                using (var package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        errors.Add("Excel file has no worksheets");
                        return (list, errors);
                    }

                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null)
                    {
                        errors.Add($"Worksheet '{ws.Name}' appears to be empty");
                        return (list, errors);
                    }

                    int maxRow = ws.Dimension.End.Row;
                    if (ws.Dimension.End.Column < 5)
                    {
                        errors.Add($"Expected at least 5 columns, found {ws.Dimension.End.Column}");
                        return (list, errors);
                    }

                    int row = 2;
                    int emptyRowCount = 0;

                    while (row <= maxRow && row <= MaxExcelRows)
                    {
                        var keyName = ws.Cells[row, 1].Text?.Trim() ?? "";
                        if (string.IsNullOrWhiteSpace(keyName))
                        {
                            emptyRowCount++;
                            if (emptyRowCount >= MaxConsecutiveEmptyRows) break;
                            row++;
                            continue;
                        }

                        emptyRowCount = 0;

                        var itemNum = ws.Cells[row, 2].Text?.Trim() ?? "";
                        var desc = ws.Cells[row, 3].Text?.Trim() ?? "";
                        var sub = ws.Cells[row, 5].Text?.Trim() ?? "";
                        double price = ParsePriceCell(ws.Cells[row, 4]);

                        list.Add(new KeyEntry
                        {
                            KeyName = keyName,
                            ItemNumber = itemNum,
                            Description = desc,
                            Price = price,
                            Subcontractor = sub
                        });

                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Error reading Excel: {ex.Message}");
                Log.Error(ex, "Error reading Excel file for Seifei Choze");
            }

            return (list, errors);
        }

        private double ParsePriceCell(ExcelRange cell)
        {
            if (cell.Value == null) return 0;
            if (cell.Value is double d) return d;
            if (cell.Value is int i) return i;
            if (cell.Value is long l) return l;
            if (cell.Value is decimal dec) return (double)dec;
            if (cell.Value is float f) return f;

            string text = cell.Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(text)) return 0;

            text = text
                .Replace("\u20AA", "").Replace("$", "").Replace("\u20AC", "")
                .Replace(",", "").Replace(" ", "").Trim();

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                return result;
            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                return result;

            return 0;
        }

        #endregion

        #region Helpers

        private string? GetExcelFilePath()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx",
                Title = "Select Key Schedule Excel File",
                CheckFileExists = true
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        private void SetParamSafe(Element e, string name, object value)
        {
            if (e == null) return;
            var p = e.LookupParameter(name);
            if (p == null || p.IsReadOnly) return;

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        if (value is double dv)
                            p.Set(dv);
                        else if (double.TryParse(value?.ToString(), NumberStyles.Any,
                                 CultureInfo.InvariantCulture, out double parsed))
                            p.Set(parsed);
                        break;

                    case StorageType.Integer:
                        if (value is int iv)
                            p.Set(iv);
                        else if (int.TryParse(value?.ToString(), out int parsedInt))
                            p.Set(parsedInt);
                        break;

                    case StorageType.String:
                        p.Set(value?.ToString() ?? "");
                        break;

                    case StorageType.ElementId:
                        if (value is ElementId eid)
                            p.Set(eid);
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Failed to set parameter '{Param}' on element {Id}", name, e.Id);
            }
        }

        private void ReportResults(int totalRows, int successCount, List<string> failures, List<string> readErrors)
        {
            string msg;
            if (failures.Count == 0 && readErrors.Count == 0)
            {
                msg = $"Success!\n\n" +
                      $"Imported {totalRows} contract items.\n" +
                      $"Updated {successCount} category key schedules.";
            }
            else
            {
                msg = $"Completed with {successCount}/{TargetCategories.Length} categories.\n" +
                      $"Imported {totalRows} contract items.\n";

                if (failures.Any())
                {
                    msg += $"\nCategory issues ({failures.Count}):\n";
                    foreach (var f in failures.Take(5))
                        msg += $"  - {f}\n";
                    if (failures.Count > 5)
                        msg += $"  ... and {failures.Count - 5} more\n";
                }

                if (readErrors.Any())
                {
                    msg += $"\nExcel notes:\n";
                    foreach (var e in readErrors)
                        msg += $"  - {e}\n";
                }
            }

            TaskDialog.Show("BimTasksV2 - Import Results", msg);
        }

        #endregion

        #region Data Classes

        private class KeyEntry
        {
            public string KeyName { get; set; } = string.Empty;
            public string ItemNumber { get; set; } = string.Empty;
            public string Description { get; set; } = string.Empty;
            public string Subcontractor { get; set; } = string.Empty;
            public double Price { get; set; }
        }

        #endregion
    }
}
