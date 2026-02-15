using System;
using System.Collections.Generic;
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
    /// Imports key schedules from an Excel file using EPPlus (OfficeOpenXml.ExcelPackage).
    /// Creates or updates key schedules for structural categories.
    /// </summary>
    public class ImportKeySchedulesFromExcelHandler : ICommandHandler
    {
        private readonly BuiltInCategory[] TargetCategories = new[]
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming
        };

        private readonly Dictionary<string, string> ParameterMapping = new Dictionary<string, string>
        {
            { "ItemNumber",    "\u05E1\u05E2\u05D9\u05E3 \u05EA\u05E7\u05E6\u05D9\u05D1\u05D9" }, // "סעיף תקציבי"
            { "Description",   "\u05EA\u05D0\u05D5\u05E8 \u05E1\u05E2\u05D9\u05E3" },               // "תאור סעיף"
            { "Subcontractor", "\u05E7\u05D1\u05DC\u05DF \u05DE\u05E9\u05E0\u05D4" },               // "קבלן משנה"
            { "Price",         "\u05DE\u05D7\u05D9\u05E8 \u05DB\u05D5\u05DC\u05DC" }                // "מחיר כולל"
        };

        public void Execute(UIApplication uiApp)
        {
            var uidoc = uiApp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // 1. Get Excel file
                string? excelPath = GetExcelFilePath();
                if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
                    return;

                // 2. Read data
                var excelData = ReadExcelData(excelPath!);
                if (excelData.Count == 0)
                {
                    TaskDialog.Show("BimTasksV2", "Excel file is empty.");
                    return;
                }

                Log.Information("ImportKeySchedules: Read {Count} entries from Excel", excelData.Count);

                // 3. Create/update key schedules
                using (var tx = new Transaction(doc, "Import Key Schedules"))
                {
                    tx.Start();

                    foreach (var bic in TargetCategories)
                    {
                        try
                        {
                            var schedule = GetOrCreateKeySchedule(doc, bic);
                            if (schedule == null) continue;

                            UpdateScheduleRows(doc, schedule, excelData);
                        }
                        catch (Exception ex)
                        {
                            Log.Warning(ex, "Error processing category {Category}", bic);
                        }
                    }

                    tx.Commit();
                }

                TaskDialog.Show("BimTasksV2", "Key Schedules updated successfully.");
                Log.Information("ImportKeySchedules: Completed successfully");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ImportKeySchedules failed");
                TaskDialog.Show("BimTasksV2", $"Error:\n{ex.Message}");
            }
        }

        #region Key Schedule Operations

        private ViewSchedule? GetOrCreateKeySchedule(Document doc, BuiltInCategory bic)
        {
            string catName = LabelUtils.GetLabelFor(bic);
            string scheduleName = $"BIMTasks_Key_{catName}";

            var schedule = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .FirstOrDefault(v => v.Name.Equals(scheduleName) && v.Definition.IsKeySchedule);

            if (schedule == null)
            {
                try
                {
                    var categoryId = Category.GetCategory(doc, bic)?.Id;
                    if (categoryId == null) return null;
                    schedule = ViewSchedule.CreateKeySchedule(doc, categoryId);
                    schedule.Name = scheduleName;
                }
                catch
                {
                    return null;
                }
            }

            // Ensure parameter fields
            foreach (var paramName in ParameterMapping.Values)
                AddFieldToSchedule(doc, schedule, paramName);

            return schedule;
        }

        private void AddFieldToSchedule(Document doc, ViewSchedule schedule, string paramName)
        {
            var definition = schedule.Definition;
            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                if (definition.GetField(i).GetName() == paramName) return;
            }

            var field = definition.GetSchedulableFields()
                .FirstOrDefault(sf => sf.GetName(doc) == paramName);

            if (field != null)
                definition.AddField(field);
        }

        private void UpdateScheduleRows(Document doc, ViewSchedule schedule, List<KeyEntry> data)
        {
            var collector = new FilteredElementCollector(doc, schedule.Id).ToElements();
            var tableData = schedule.GetTableData();
            var sectionData = tableData.GetSectionData(SectionType.Body);

            foreach (var entry in data)
            {
                var keyElement = collector.FirstOrDefault(e => e.Name == entry.KeyName);

                if (keyElement == null)
                {
                    sectionData.InsertRow(sectionData.FirstRowNumber);

                    var newCollector = new FilteredElementCollector(doc, schedule.Id).ToElements();
                    keyElement = newCollector
                        .Except(collector, new ElementIdEqualityComparer())
                        .FirstOrDefault();

                    if (keyElement != null)
                    {
                        var updatedList = collector.ToList();
                        updatedList.Add(keyElement);
                        collector = updatedList;
                    }
                }

                if (keyElement != null)
                {
                    SetParam(keyElement, "Key Name", entry.KeyName);
                    SetParam(keyElement, ParameterMapping["ItemNumber"], entry.ItemNumber);
                    SetParam(keyElement, ParameterMapping["Description"], entry.Description);
                    SetParam(keyElement, ParameterMapping["Subcontractor"], entry.Subcontractor);
                    SetParam(keyElement, ParameterMapping["Price"], entry.Price);
                }
            }
        }

        #endregion

        #region Helpers

        private void SetParam(Element e, string name, object value)
        {
            var p = e.LookupParameter(name);
            if (p != null && !p.IsReadOnly)
            {
                if (value is double d) p.Set(d);
                else if (value is string s) p.Set(s);
            }
        }

        private List<KeyEntry> ReadExcelData(string path)
        {
            var list = new List<KeyEntry>();

            try
            {
                byte[] fileBytes = File.ReadAllBytes(path);
                using (var stream = new MemoryStream(fileBytes))
                using (var package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0) return list;

                    var ws = package.Workbook.Worksheets[0];
                    if (ws.Dimension == null) return list;

                    int row = 2;
                    while (row <= ws.Dimension.End.Row)
                    {
                        var keyName = ws.Cells[row, 1].Text?.Trim() ?? "";
                        if (string.IsNullOrWhiteSpace(keyName)) break;

                        var itemNum = ws.Cells[row, 2].Text?.Trim() ?? "";
                        var desc = ws.Cells[row, 3].Text?.Trim() ?? "";
                        var priceText = ws.Cells[row, 4].Text?.Trim() ?? "";
                        var sub = ws.Cells[row, 5].Text?.Trim() ?? "";

                        double.TryParse(priceText, out double price);

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
                Log.Error(ex, "Error reading Excel for key schedules");
            }

            return list;
        }

        private string? GetExcelFilePath()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                Title = "Select Key Schedule Excel File"
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
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

        private class ElementIdEqualityComparer : IEqualityComparer<Element>
        {
            public bool Equals(Element? x, Element? y)
            {
                if (x == null || y == null) return false;
                return x.Id == y.Id;
            }

            public int GetHashCode(Element obj) => obj.Id.GetHashCode();
        }

        #endregion
    }
}
