using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Serilog;
using XlColor = DocumentFormat.OpenXml.Spreadsheet.Color;
using XlFont = DocumentFormat.OpenXml.Spreadsheet.Font;
using XlBorder = DocumentFormat.OpenXml.Spreadsheet.Border;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Exports Revit schedules to Excel and imports changes back using OpenXml.
    /// </summary>
    public class ScheduleExcelRoundtripService
    {
        private const int ElementIdColumn = 1;
        private const int DataStartColumn = 2;
        private const int HeaderRow = 1;
        private const int DataStartRow = 2;
        private const string ElementIdHeader = "__ElementIds__";

        // Style indices
        private const uint STL_DEFAULT = 0;
        private const uint STL_HEADER = 1;
        private const uint STL_NORMAL = 2;
        private const uint STL_READONLY = 3;
        private const uint STL_HEADER_RO = 4;

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

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = spreadsheet.WorkbookPart;
                if (workbookPart == null)
                {
                    result.Errors.Add("Invalid Excel file: no workbook part found.");
                    return result;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var sst = workbookPart.SharedStringTablePart?.SharedStringTable;

                var allRows = sheetData.Elements<Row>().ToList();
                if (allRows.Count < DataStartRow)
                {
                    Log.Warning("Excel file has no data rows");
                    return result;
                }

                // Build header-to-field mapping
                var headerRowData = allRows.FirstOrDefault(r => r.RowIndex != null && r.RowIndex == (uint)HeaderRow);
                var colToField = new Dictionary<int, FieldInfo>();
                if (headerRowData != null)
                {
                    foreach (var cell in headerRowData.Elements<Cell>())
                    {
                        int colIdx = GetColumnIndex(cell.CellReference);
                        if (colIdx < DataStartColumn) continue;

                        string? header = GetCellValue(cell, sst)?.Trim();
                        if (header != null && header.EndsWith(" (read-only)"))
                            header = header.Substring(0, header.Length - " (read-only)".Length);

                        var field = fields.FirstOrDefault(f => f.Name == header);
                        if (field != null && field.IsWritable)
                            colToField[colIdx] = field;
                    }
                }

                if (colToField.Count == 0)
                {
                    Log.Warning("No writable columns found in Excel file");
                    return result;
                }

                // Process data rows
                foreach (var row in allRows.Where(r => r.RowIndex != null && r.RowIndex >= (uint)DataStartRow))
                {
                    var cellMap = new Dictionary<int, string>();
                    foreach (var cell in row.Elements<Cell>())
                    {
                        int colIdx = GetColumnIndex(cell.CellReference);
                        cellMap[colIdx] = GetCellValue(cell, sst) ?? "";
                    }

                    if (!cellMap.TryGetValue(ElementIdColumn, out string? elementIdsText) ||
                        string.IsNullOrEmpty(elementIdsText?.Trim()))
                        continue;

                    var elementIds = ParseElementIds(elementIdsText!);
                    if (elementIds.Count == 0) continue;

                    foreach (var kvp in colToField)
                    {
                        int col = kvp.Key;
                        var fieldInfo = kvp.Value;
                        cellMap.TryGetValue(col, out string? excelValue);
                        excelValue = excelValue?.Trim() ?? "";

                        foreach (var elemId in elementIds)
                        {
                            var element = doc.GetElement(elemId);
                            if (element == null)
                            {
                                result.FailedCells++;
                                result.Errors.Add($"Row {row.RowIndex}: Element {elemId.Value} not found");
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
                                result.Errors.Add($"Row {row.RowIndex}, '{fieldInfo.Name}', Element {elemId.Value}: {ex.Message}");
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

            using (var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Stylesheet
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = BuildRoundtripStylesheet();
                stylesPart.Stylesheet.Save();

                // Worksheet
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet();

                // Freeze header row
                worksheet.Append(new SheetViews(
                    new SheetView(
                        new Pane
                        {
                            VerticalSplit = 1D,
                            TopLeftCell = "A2",
                            ActivePane = PaneValues.BottomLeft,
                            State = PaneStateValues.Frozen
                        },
                        new Selection { Pane = PaneValues.BottomLeft }
                    )
                    {
                        TabSelected = true,
                        WorkbookViewId = 0U
                    }
                ));

                // Column widths
                var cols = new Columns();
                cols.Append(new Column { Min = 1, Max = 1, Width = 12, Hidden = true, CustomWidth = true });
                for (int i = 0; i < fields.Count; i++)
                {
                    double maxLen = fields[i].Name.Length;
                    foreach (var row in rows)
                    {
                        if (i < row.Values.Count)
                            maxLen = Math.Max(maxLen, (row.Values[i] ?? "").Length);
                    }
                    double width = Math.Max(10, Math.Min(50, maxLen * 1.2 + 2));
                    uint colIdx = (uint)(DataStartColumn + i);
                    cols.Append(new Column { Min = colIdx, Max = colIdx, Width = width, CustomWidth = true });
                }
                worksheet.Append(cols);

                // Sheet data
                var sheetData = new SheetData();

                // Header row
                var headerRow = new Row { RowIndex = (uint)HeaderRow };
                headerRow.Append(MakeStringCell(HeaderRow, ElementIdColumn, ElementIdHeader, STL_HEADER));
                foreach (var fi in fields)
                {
                    string displayName = fi.IsWritable ? fi.Name : fi.Name + " (read-only)";
                    uint style = fi.IsWritable ? STL_HEADER : STL_HEADER_RO;
                    headerRow.Append(MakeStringCell(HeaderRow, fi.ExcelColumn, displayName, style));
                }
                sheetData.Append(headerRow);

                // Data rows
                for (int r = 0; r < rows.Count; r++)
                {
                    int excelRow = DataStartRow + r;
                    var row = rows[r];
                    var dataRow = new Row { RowIndex = (uint)excelRow };

                    string idsText = string.Join(",", row.ElementIds.Select(id => id.Value));
                    dataRow.Append(MakeStringCell(excelRow, ElementIdColumn, idsText, STL_DEFAULT));

                    for (int f = 0; f < fields.Count; f++)
                    {
                        uint style = fields[f].IsWritable ? STL_NORMAL : STL_READONLY;
                        dataRow.Append(MakeStringCell(excelRow, fields[f].ExcelColumn, row.Values[f], style));
                    }

                    sheetData.Append(dataRow);
                }

                worksheet.Append(sheetData);

                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                workbookPart.Workbook.Append(new Sheets(
                    new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1U,
                        Name = safeSheet
                    }
                ));

                workbookPart.Workbook.Save();
            }
        }

        #endregion

        #region OpenXML Helpers

        private static Stylesheet BuildRoundtripStylesheet()
        {
            var fonts = new Fonts(
                new XlFont(new FontSize { Val = 10 }, new FontName { Val = "Calibri" }),
                new XlFont(new Bold(), new FontSize { Val = 10 }, new XlColor { Rgb = "FFFFFFFF" },
                    new FontName { Val = "Calibri" }),
                new XlFont(new FontSize { Val = 10 }, new XlColor { Rgb = "FF808080" },
                    new FontName { Val = "Calibri" })
            ) { Count = 3 };

            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                new Fill(new PatternFill(new ForegroundColor { Rgb = "FF4472C4" }) { PatternType = PatternValues.Solid }),
                new Fill(new PatternFill(new ForegroundColor { Rgb = "FFE6E6E6" }) { PatternType = PatternValues.Solid })
            ) { Count = 4 };

            var borders = new Borders(
                new XlBorder(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
                new XlBorder(
                    new LeftBorder(new XlColor { Rgb = "FFD0D0D0" }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new XlColor { Rgb = "FFD0D0D0" }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new XlColor { Rgb = "FFD0D0D0" }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new XlColor { Rgb = "FFD0D0D0" }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder())
            ) { Count = 2 };

            var cellFormats = new CellFormats(
                new CellFormat(),
                new CellFormat
                {
                    FontId = 1, FillId = 2, BorderId = 1,
                    ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                    ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center }
                },
                new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyFont = true, ApplyBorder = true },
                new CellFormat { FontId = 2, FillId = 3, BorderId = 1, ApplyFont = true, ApplyFill = true, ApplyBorder = true },
                new CellFormat
                {
                    FontId = 1, FillId = 3, BorderId = 1,
                    ApplyFont = true, ApplyFill = true, ApplyBorder = true,
                    ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center }
                }
            ) { Count = 5 };

            return new Stylesheet(fonts, fills, borders, cellFormats);
        }

        private static Cell MakeStringCell(int row, int col, string value, uint styleIndex)
        {
            return new Cell
            {
                CellReference = $"{ColLetter(col)}{row}",
                DataType = CellValues.InlineString,
                StyleIndex = styleIndex,
                InlineString = new InlineString(new Text(value ?? "") { Space = SpaceProcessingModeValues.Preserve })
            };
        }

        private static string ColLetter(int col)
        {
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + col % 26) + result;
                col /= 26;
            }
            return result;
        }

        private static int GetColumnIndex(string? cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;
            int index = 0;
            foreach (char c in cellReference)
            {
                if (!char.IsLetter(c)) break;
                index = index * 26 + (char.ToUpper(c) - 'A' + 1);
            }
            return index;
        }

        private static string GetCellValue(Cell cell, SharedStringTable? sst)
        {
            if (cell == null) return "";

            if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? "";

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (cell.CellValue != null && int.TryParse(cell.CellValue.InnerText, out int idx) && sst != null)
                    return sst.ElementAt(idx).InnerText;
                return "";
            }

            return cell.CellValue?.InnerText ?? "";
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
