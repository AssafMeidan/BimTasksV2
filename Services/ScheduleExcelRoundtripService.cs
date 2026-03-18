using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Exports Revit schedules to Excel and imports changes back using ClosedXML.
    /// </summary>
    public class ScheduleExcelRoundtripService
    {
        private const int ElementIdColumn = 1;
        private const int DataStartColumn = 2;
        private const int HeaderRow = 1;
        private const int DataStartRow = 2;
        private const string ElementIdHeader = "__ElementIds__";

        /// <summary>
        /// Metadata about a schedule field used during export/import.
        /// </summary>
        private class FieldInfo
        {
            public ScheduleField Field { get; set; } = null!;
            public string Name { get; set; } = string.Empty;
            public bool IsWritable { get; set; }
            public bool IsTypeParameter { get; set; }
            public int ExcelColumn { get; set; }
        }

        /// <summary>
        /// Exports a schedule to an Excel file with element tracking.
        /// </summary>
        public void ExportScheduleToExcel(ViewSchedule schedule, string filePath)
        {
            var doc = schedule.Document;
            var definition = schedule.Definition;
            var fields = AnalyzeFields(doc, schedule);

            if (fields.Count == 0)
                throw new InvalidOperationException("Schedule has no exportable fields.");

            var rows = BuildExportRows(doc, schedule, definition, fields);

            WriteExcelFile(filePath, schedule.Name, fields, rows);

            Log.Information("Exported schedule '{ScheduleName}' to {FilePath}: {RowCount} rows, {FieldCount} fields",
                schedule.Name, filePath, rows.Count, fields.Count);
        }

        /// <summary>
        /// Imports changes from an edited Excel file back into the Revit model.
        /// Returns ImportResult with cell-level update counts.
        /// </summary>
        public Models.ImportResult ImportScheduleFromExcel(string filePath, Document doc, ViewSchedule schedule)
        {
            var result = new Models.ImportResult
            {
                ScheduleName = schedule.Name
            };
            var fields = AnalyzeFields(doc, schedule);
            var affectedElementIds = new HashSet<ElementId>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var ws = workbook.Worksheets.First();
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;

                if (lastRow < DataStartRow)
                {
                    Log.Warning("Excel file has no data rows");
                    return result;
                }

                // Build header-to-field mapping
                var colToField = new Dictionary<int, FieldInfo>();
                for (int col = DataStartColumn; col <= lastCol; col++)
                {
                    string header = ws.Cell(HeaderRow, col).GetString().Trim();
                    // Support legacy files that had " (read-only)" suffix
                    if (header.EndsWith(" (read-only)"))
                        header = header.Substring(0, header.Length - " (read-only)".Length);

                    var field = fields.FirstOrDefault(f => f.Name == header);
                    if (field != null && field.IsWritable)
                        colToField[col] = field;
                }

                if (colToField.Count == 0)
                {
                    Log.Warning("No writable columns found in Excel file");
                    return result;
                }

                // Build lookup table if overflow sheet exists
                var groupLookup = new Dictionary<string, List<ElementId>>();
                if (workbook.Worksheets.TryGetWorksheet("__Lookup__", out var lookupWs))
                {
                    int lookupLastRow = lookupWs.LastRowUsed()?.RowNumber() ?? 0;
                    for (int lr = 1; lr <= lookupLastRow; lr++)
                    {
                        string key = lookupWs.Cell(lr, 1).GetString().Trim();
                        if (string.IsNullOrEmpty(key)) continue;
                        if (long.TryParse(lookupWs.Cell(lr, 2).GetString().Trim(), out long id))
                        {
                            if (!groupLookup.ContainsKey(key))
                                groupLookup[key] = new List<ElementId>();
                            groupLookup[key].Add(new ElementId(id));
                        }
                    }
                }

                // Process data rows
                for (int row = DataStartRow; row <= lastRow; row++)
                {
                    string elementIdsText = ws.Cell(row, ElementIdColumn).GetString().Trim();
                    if (string.IsNullOrEmpty(elementIdsText))
                        continue;

                    // Resolve group key or parse IDs directly
                    List<ElementId> elementIds;
                    if (elementIdsText.StartsWith("__G") && elementIdsText.EndsWith("__") &&
                        groupLookup.TryGetValue(elementIdsText, out var lookupIds))
                    {
                        elementIds = lookupIds;
                    }
                    else
                    {
                        elementIds = ParseElementIds(elementIdsText);
                    }
                    if (elementIds.Count == 0) continue;

                    foreach (var kvp in colToField)
                    {
                        int col = kvp.Key;
                        var fieldInfo = kvp.Value;
                        string excelValue = ws.Cell(row, col).GetString().Trim();

                        foreach (var elemId in elementIds)
                        {
                            var element = doc.GetElement(elemId);
                            if (element == null)
                            {
                                result.FailedCells++;
                                result.Errors.Add($"Row {row}: Element {elemId.Value} not found");
                                continue;
                            }

                            try
                            {
                                bool changed = SetParameterValue(doc, element, fieldInfo, excelValue);
                                if (changed)
                                {
                                    result.UpdatedCells++;
                                    affectedElementIds.Add(elemId);
                                }
                                else
                                {
                                    result.SkippedCells++;
                                }
                            }
                            catch (Exception ex)
                            {
                                result.FailedCells++;
                                result.Errors.Add($"Row {row}, '{fieldInfo.Name}', Element {elemId.Value}: {ex.Message}");
                            }
                        }
                    }

                    result.RowsImported++;
                }
            }

            result.AffectedElements = affectedElementIds.Count;

            Log.Information("Imported schedule changes: {Updated} updated, {Skipped} skipped, {Failed} failed, {Elements} elements affected",
                result.UpdatedCells, result.SkippedCells, result.FailedCells, result.AffectedElements);

            if (result.Errors.Count > 0)
                LogFailuresToFile(result.Errors, schedule.Name);

            return result;
        }

        #region Field Analysis

        private List<FieldInfo> AnalyzeFields(Document doc, ViewSchedule schedule)
        {
            var definition = schedule.Definition;
            var fields = new List<FieldInfo>();
            int colIndex = DataStartColumn;

            for (int i = 0; i < definition.GetFieldCount(); i++)
            {
                var field = definition.GetField(i);
                if (field.IsHidden)
                    continue;

                bool isWritable = false;
                bool isTypeParam = false;

                if (field.FieldType == ScheduleFieldType.Instance ||
                    field.FieldType == ScheduleFieldType.ElementType)
                {
                    isTypeParam = field.FieldType == ScheduleFieldType.ElementType;

                    var sampleElements = new FilteredElementCollector(doc, schedule.Id)
                        .WhereElementIsNotElementType()
                        .Take(1)
                        .ToList();

                    if (sampleElements.Count > 0)
                    {
                        var param = GetParameter(doc, sampleElements[0], field);
                        if (param != null && !param.IsReadOnly)
                            isWritable = true;
                    }
                }

                fields.Add(new FieldInfo
                {
                    Field = field,
                    Name = field.GetName(),
                    IsWritable = isWritable,
                    IsTypeParameter = isTypeParam,
                    ExcelColumn = colIndex
                });

                colIndex++;
            }

            return fields;
        }

        #endregion

        #region Export Rows

        private List<ExportRow> BuildExportRows(Document doc, ViewSchedule schedule,
            ScheduleDefinition definition, List<FieldInfo> fields)
        {
            var rows = new List<ExportRow>();

            var tableData = schedule.GetTableData();
            var bodyData = tableData.GetSectionData(SectionType.Body);
            int bodyRowCount = bodyData.NumberOfRows;
            int bodyColCount = bodyData.NumberOfColumns;

            var allElements = new FilteredElementCollector(doc, schedule.Id)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            if (definition.IsItemized)
            {
                var elementSignatures = new Dictionary<string, List<Element>>();
                foreach (var elem in allElements)
                {
                    string sig = BuildElementSignature(doc, elem, fields);
                    if (!elementSignatures.ContainsKey(sig))
                        elementSignatures[sig] = new List<Element>();
                    elementSignatures[sig].Add(elem);
                }

                var assignedElementIds = new HashSet<ElementId>();

                for (int row = 0; row < bodyRowCount; row++)
                {
                    string rowSig = BuildRowSignature(schedule, row, bodyColCount);
                    if (string.IsNullOrWhiteSpace(rowSig)) continue;

                    Element? matchedElement = null;
                    if (elementSignatures.TryGetValue(rowSig, out var candidates))
                        matchedElement = candidates.FirstOrDefault(e => !assignedElementIds.Contains(e.Id));

                    if (matchedElement == null)
                    {
                        if (IsLikelyTotalRow(schedule, row, bodyColCount, fields.Count))
                            continue;
                        continue;
                    }

                    assignedElementIds.Add(matchedElement.Id);

                    var cellValues = new List<string>();
                    foreach (var fi in fields)
                        cellValues.Add(GetParameterDisplayValue(doc, matchedElement, fi));

                    rows.Add(new ExportRow
                    {
                        ElementIds = new List<ElementId> { matchedElement.Id },
                        Values = cellValues
                    });
                }

                foreach (var elem in allElements.Where(e => !assignedElementIds.Contains(e.Id)))
                {
                    var cellValues = new List<string>();
                    foreach (var fi in fields)
                        cellValues.Add(GetParameterDisplayValue(doc, elem, fi));
                    rows.Add(new ExportRow
                    {
                        ElementIds = new List<ElementId> { elem.Id },
                        Values = cellValues
                    });
                }
            }
            else
            {
                var groups = new Dictionary<string, List<Element>>();
                foreach (var elem in allElements)
                {
                    string key = BuildElementSignature(doc, elem, fields);
                    if (!groups.ContainsKey(key))
                        groups[key] = new List<Element>();
                    groups[key].Add(elem);
                }

                foreach (var group in groups)
                {
                    var representative = group.Value[0];
                    var cellValues = new List<string>();
                    foreach (var fi in fields)
                        cellValues.Add(GetParameterDisplayValue(doc, representative, fi));

                    rows.Add(new ExportRow
                    {
                        ElementIds = group.Value.Select(e => e.Id).ToList(),
                        Values = cellValues
                    });
                }
            }

            return rows;
        }

        private string BuildElementSignature(Document doc, Element element, List<FieldInfo> fields)
        {
            var parts = new List<string>();
            foreach (var fi in fields)
                parts.Add(GetParameterDisplayValue(doc, element, fi) ?? "");
            return string.Join("|", parts);
        }

        private string BuildRowSignature(ViewSchedule schedule, int row, int colCount)
        {
            var parts = new List<string>();
            for (int col = 0; col < colCount; col++)
                parts.Add(schedule.GetCellText(SectionType.Body, row, col) ?? "");
            return string.Join("|", parts);
        }

        private bool IsLikelyTotalRow(ViewSchedule schedule, int row, int colCount, int fieldCount)
        {
            int filledCells = 0;
            int emptyCells = 0;

            for (int col = 0; col < colCount; col++)
            {
                string text = schedule.GetCellText(SectionType.Body, row, col);
                if (string.IsNullOrWhiteSpace(text))
                    emptyCells++;
                else
                    filledCells++;
            }

            if (fieldCount > 2 && emptyCells > filledCells)
                return true;

            if (filledCells == 1)
            {
                for (int col = 0; col < colCount; col++)
                {
                    string text = schedule.GetCellText(SectionType.Body, row, col);
                    if (!string.IsNullOrEmpty(text) && (text.Contains(":") || text.Contains("Grand")))
                        return true;
                }
            }

            return false;
        }

        #endregion

        #region Parameter Helpers

        private string GetParameterDisplayValue(Document doc, Element element, FieldInfo fieldInfo)
        {
            var param = GetParameter(doc, element, fieldInfo.Field);
            if (param == null) return "";

            if (param.StorageType == StorageType.String)
                return param.AsString() ?? "";

            string? valueStr = param.AsValueString();
            return valueStr ?? "";
        }

        private Autodesk.Revit.DB.Parameter? GetParameter(Document doc, Element element, ScheduleField field)
        {
            Element target = element;
            if (field.FieldType == ScheduleFieldType.ElementType)
            {
                var typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                    return null;
                target = doc.GetElement(typeId);
                if (target == null) return null;
            }

            var paramId = field.ParameterId;
            if (paramId == null || paramId == ElementId.InvalidElementId)
                return null;

            if (paramId.Value < 0)
            {
                var bip = (BuiltInParameter)paramId.Value;
                return target.get_Parameter(bip);
            }

            var paramElem = doc.GetElement(paramId);
            if (paramElem != null)
            {
                string paramName = paramElem.Name;
                return target.LookupParameter(paramName);
            }

            return null;
        }

        private bool SetParameterValue(Document doc, Element element, FieldInfo fieldInfo, string newValue)
        {
            Element target = element;
            if (fieldInfo.IsTypeParameter)
            {
                var typeId = element.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                    return false;
                target = doc.GetElement(typeId);
                if (target == null) return false;
            }

            var param = GetParameter(doc, element, fieldInfo.Field);
            if (param == null || param.IsReadOnly)
                return false;

            string currentValue = param.StorageType == StorageType.String
                ? (param.AsString() ?? "")
                : (param.AsValueString() ?? "");

            if (currentValue == (newValue ?? ""))
                return false;

            switch (param.StorageType)
            {
                case StorageType.String:
                    param.Set(newValue ?? "");
                    return true;

                case StorageType.Integer:
                    if (int.TryParse(newValue, out int intVal))
                    {
                        param.Set(intVal);
                        return true;
                    }
                    return param.SetValueString(newValue);

                case StorageType.Double:
                    return param.SetValueString(newValue);

                case StorageType.ElementId:
                    return param.SetValueString(newValue);

                default:
                    return false;
            }
        }

        #endregion

        #region Excel Writing

        private void WriteExcelFile(string filePath, string sheetName, List<FieldInfo> fields, List<ExportRow> rows)
        {
            string safeSheet = string.Join("_",
                sheetName.Split(new[] { '\\', '/', '?', '*', '[', ']', ':' }));
            if (safeSheet.Length > 31)
                safeSheet = safeSheet.Substring(0, 31);

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add(safeSheet);

                // --- Header row ---
                ws.Cell(HeaderRow, ElementIdColumn).Value = ElementIdHeader;

                foreach (var fi in fields)
                    ws.Cell(HeaderRow, fi.ExcelColumn).Value = fi.Name;

                // --- Overflow IDs sheet (created only if needed) ---
                const int MaxCellChars = 32000; // Excel limit is 32767, leave margin
                IXLWorksheet? lookupSheet = null;
                int lookupRow = 1;

                // --- Data rows ---
                for (int r = 0; r < rows.Count; r++)
                {
                    int excelRow = DataStartRow + r;
                    var row = rows[r];

                    string idsText = string.Join(",", row.ElementIds.Select(id => id.Value));

                    if (idsText.Length <= MaxCellChars)
                    {
                        ws.Cell(excelRow, ElementIdColumn).Value = idsText;
                    }
                    else
                    {
                        // Store a group key in main sheet, full IDs in lookup sheet
                        string groupKey = $"__G{r}__";
                        ws.Cell(excelRow, ElementIdColumn).Value = groupKey;

                        if (lookupSheet == null)
                        {
                            lookupSheet = workbook.Worksheets.Add("__Lookup__");
                            lookupSheet.Visibility = XLWorksheetVisibility.VeryHidden;
                        }

                        foreach (var id in row.ElementIds)
                        {
                            lookupSheet.Cell(lookupRow, 1).Value = groupKey;
                            lookupSheet.Cell(lookupRow, 2).Value = id.Value;
                            lookupRow++;
                        }
                    }

                    for (int f = 0; f < fields.Count; f++)
                        ws.Cell(excelRow, fields[f].ExcelColumn).Value = row.Values[f] ?? "";
                }

                // Hide ElementIds column
                ws.Column(ElementIdColumn).Hide();

                workbook.SaveAs(filePath);
            }
        }


        #endregion

        #region Misc Helpers

        private void LogFailuresToFile(List<string> failures, string scheduleName)
        {
            try
            {
                string logFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "BimTasksV2", "Logs");
                if (!Directory.Exists(logFolder))
                    Directory.CreateDirectory(logFolder);

                string logFile = Path.Combine(logFolder,
                    $"ScheduleImport_{scheduleName}_{DateTime.Now:yyyyMMdd_HHmmss}.log");

                File.WriteAllLines(logFile, failures);
                Log.Information("Import failures logged to {LogFile}", logFile);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to write import failure log");
            }
        }

        private List<ElementId> ParseElementIds(string text)
        {
            var ids = new List<ElementId>();
            foreach (var part in text.Split(','))
            {
                if (long.TryParse(part.Trim(), out long id))
                    ids.Add(new ElementId(id));
            }
            return ids;
        }

        private class ExportRow
        {
            public List<ElementId> ElementIds { get; set; } = new();
            public List<string> Values { get; set; } = new();
        }

        #endregion
    }
}
