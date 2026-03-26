using System;
using System.Collections.Generic;
using System.Globalization;
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
    /// Professional one-way schedule-to-Excel export with full formatting.
    /// Ported from old BIMTasks ExportSchedualeToExcelCommand.
    /// </summary>
    public class ScheduleExcelExportService
    {
        // ── Style indices (27 CellFormats) ──
        private const uint STL_DEFAULT = 0;
        private const uint STL_HEADER = 1;
        // White (even) rows
        private const uint STL_W_INT = 2;
        private const uint STL_W_DEC_ZERO = 3;
        private const uint STL_W_DEC = 4;
        private const uint STL_W_TEXT = 5;
        private const uint STL_W_PCT = 6;
        private const uint STL_W_CENTER_WRAP = 7;
        // Alt (odd) rows
        private const uint STL_A_INT = 8;
        private const uint STL_A_DEC = 9;
        private const uint STL_A_TEXT = 10;
        private const uint STL_A_PCT = 11;
        private const uint STL_A_CENTER_WRAP = 12;
        // Subtotal
        private const uint STL_SUB_INT = 13;
        private const uint STL_SUB_DEC = 14;
        private const uint STL_SUB_TEXT = 15;
        private const uint STL_SUB_CENTER = 16;
        private const uint STL_SUB_CENTER_WRAP = 17;
        // Alt decimal zero (gray font)
        private const uint STL_A_DEC_ZERO = 18;
        // Grand total
        private const uint STL_GT_INT = 19;
        private const uint STL_GT_DEC = 20;
        private const uint STL_GT_TEXT = 21;
        private const uint STL_GT_CENTER = 22;
        private const uint STL_GT_CENTER_WRAP = 23;
        // Title (merged)
        private const uint STL_TITLE_L = 24;
        private const uint STL_TITLE_M = 25;
        private const uint STL_TITLE_R = 26;

        private enum ColumnType { Text, Int, Dec, Pct }
        private enum CellKind { Int, Dec, DecZero, Pct, Text, Empty }

        /// <summary>
        /// Extracts schedule data and exports to a professionally formatted Excel file.
        /// </summary>
        public void ExportScheduleToExcel(ViewSchedule schedule, string filePath)
        {
            ExportSchedulesToExcel(new List<ViewSchedule> { schedule }, filePath);
        }

        /// <summary>
        /// Exports multiple schedules into a single Excel file, one sheet per schedule.
        /// </summary>
        public void ExportSchedulesToExcel(List<ViewSchedule> schedules, string filePath)
        {
            if (schedules == null || schedules.Count == 0)
                throw new ArgumentException("No schedules to export.");

            var scheduleData = new List<(string name, List<string> headers, List<List<string>> rows)>();
            foreach (var schedule in schedules)
            {
                var (headers, rows) = ExtractScheduleData(schedule);
                if (rows.Count > 0)
                    scheduleData.Add((schedule.Name, headers, rows));
            }

            if (scheduleData.Count == 0)
                throw new InvalidOperationException("The selected schedules have no data to export.");

            ExportToExcel(scheduleData, filePath);

            Log.Information("ScheduleExcelExportService: Exported {Count} schedule(s) to {FilePath}",
                scheduleData.Count, filePath);
        }

        #region Data Extraction

        private (List<string> headers, List<List<string>> rows) ExtractScheduleData(ViewSchedule schedule)
        {
            var headers = new List<string>();
            var rows = new List<List<string>>();

            TableData tableData = schedule.GetTableData();
            TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);

            int numRows = bodyData.NumberOfRows;
            int numCols = bodyData.NumberOfColumns;

            if (numRows == 0 || numCols == 0)
                return (headers, rows);

            var fields = schedule.Definition.GetFieldOrder()
                .Select(id => schedule.Definition.GetField(id))
                .Where(f => !f.IsHidden)
                .ToList();

            // Use ColumnHeading (falls back to GetName if empty)
            if (fields.Count == numCols)
            {
                headers = fields.Select(f =>
                {
                    string heading = f.ColumnHeading;
                    return string.IsNullOrEmpty(heading) ? f.GetName() : heading;
                }).ToList();
            }
            else
            {
                for (int col = 0; col < numCols; col++)
                    headers.Add(schedule.GetCellText(SectionType.Body, 0, col));
            }

            // Skip body row 0 if it matches headers (compare against both ColumnHeading and GetName)
            int startRow = 0;
            if (numRows > 1 && fields.Count == numCols)
            {
                int matchCount = 0;
                for (int col = 0; col < numCols; col++)
                {
                    string cellText = schedule.GetCellText(SectionType.Body, 0, col);
                    string heading = string.IsNullOrEmpty(fields[col].ColumnHeading)
                        ? fields[col].GetName() : fields[col].ColumnHeading;
                    string name = fields[col].GetName();
                    if (cellText == heading || cellText == name)
                        matchCount++;
                }
                if (matchCount >= numCols / 2.0)
                    startRow = 1;
            }

            for (int row = startRow; row < numRows; row++)
            {
                var rowData = new List<string>();
                for (int col = 0; col < numCols; col++)
                    rowData.Add(schedule.GetCellText(SectionType.Body, row, col));
                rows.Add(rowData);
            }

            return (headers, rows);
        }

        #endregion

        #region Row & Column Classification

        private List<int> ClassifyRows(List<List<string>> rows, int colCount)
        {
            var result = new List<int>();
            var summaryRowIndices = new List<int>();

            for (int i = 0; i < rows.Count; i++)
            {
                int emptyCount = rows[i].Count(c => string.IsNullOrWhiteSpace(c));
                bool isSummary = emptyCount > colCount / 2 && emptyCount > 0;
                if (isSummary)
                    summaryRowIndices.Add(i);
                result.Add(isSummary ? 1 : 0);
            }

            if (summaryRowIndices.Count > 0)
                result[summaryRowIndices.Last()] = 2;

            return result;
        }

        private ColumnType[] DetectColumnTypes(List<string> headers, List<List<string>> rows, List<int> rowTypes)
        {
            int colCount = headers.Count;
            var types = new ColumnType[colCount];

            for (int c = 0; c < colCount; c++)
            {
                string h = headers[c];
                if (h.Contains("%") || h.Contains("אחוז"))
                {
                    types[c] = ColumnType.Pct;
                    continue;
                }

                bool anyParsed = false;
                bool anyFractional = false;
                bool foundPct = false;
                double maxVal = 0;

                for (int r = 0; r < rows.Count; r++)
                {
                    if (rowTypes[r] != 0) continue;
                    if (c >= rows[r].Count) continue;
                    string text = (rows[r][c] ?? "").Trim();
                    if (string.IsNullOrEmpty(text)) continue;

                    if (text.EndsWith("%"))
                    {
                        string numPart = text.TrimEnd('%').Trim();
                        if (double.TryParse(numPart, NumberStyles.Any, CultureInfo.CurrentCulture, out _))
                        {
                            foundPct = true;
                            break;
                        }
                    }

                    if (TryParseNumber(text, out double val))
                    {
                        anyParsed = true;
                        if (Math.Abs(val) > maxVal) maxVal = Math.Abs(val);
                        if (val != Math.Floor(val))
                            anyFractional = true;
                    }
                }

                if (foundPct)
                    types[c] = ColumnType.Pct;
                else if (!anyParsed)
                    types[c] = ColumnType.Text;
                else if (anyFractional)
                    types[c] = ColumnType.Dec;
                else if (maxVal >= 10)
                    types[c] = ColumnType.Dec;
                else
                    types[c] = ColumnType.Int;
            }

            return types;
        }

        private static bool TryParseNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return true;

            // Find last digit/comma/dot and take substring up to there (strips unit suffixes)
            int lastNumPos = -1;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                char ch = text[i];
                if (char.IsDigit(ch) || ch == ',' || ch == '.')
                {
                    lastNumPos = i;
                    break;
                }
            }

            if (lastNumPos < 0)
                return false;

            string numPart = text.Substring(0, lastNumPos + 1).Trim();
            return double.TryParse(numPart, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
        }

        #endregion

        #region Style Selection

        private static uint GetCellStyle(int rowType, bool isAlt, CellKind kind)
        {
            if (rowType == 2) // Grand total
            {
                switch (kind)
                {
                    case CellKind.Int: return STL_GT_INT;
                    case CellKind.Dec:
                    case CellKind.DecZero:
                    case CellKind.Pct: return STL_GT_DEC;
                    case CellKind.Text: return STL_GT_TEXT;
                    default: return STL_GT_CENTER;
                }
            }
            if (rowType == 1) // Subtotal
            {
                switch (kind)
                {
                    case CellKind.Int: return STL_SUB_INT;
                    case CellKind.Dec:
                    case CellKind.DecZero:
                    case CellKind.Pct: return STL_SUB_DEC;
                    case CellKind.Text: return STL_SUB_TEXT;
                    default: return STL_SUB_CENTER;
                }
            }
            // Normal row
            if (isAlt)
            {
                switch (kind)
                {
                    case CellKind.Int: return STL_A_INT;
                    case CellKind.Dec: return STL_A_DEC;
                    case CellKind.DecZero: return STL_A_DEC_ZERO;
                    case CellKind.Pct: return STL_A_PCT;
                    case CellKind.Text: return STL_A_TEXT;
                    default: return STL_A_CENTER_WRAP;
                }
            }
            switch (kind)
            {
                case CellKind.Int: return STL_W_INT;
                case CellKind.Dec: return STL_W_DEC;
                case CellKind.DecZero: return STL_W_DEC_ZERO;
                case CellKind.Pct: return STL_W_PCT;
                case CellKind.Text: return STL_W_TEXT;
                default: return STL_W_CENTER_WRAP;
            }
        }

        #endregion

        #region Excel Export

        private void ExportToExcel(
            List<(string name, List<string> headers, List<List<string>> rows)> scheduleData,
            string filePath)
        {
            using (var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = BuildStylesheet();
                stylesPart.Stylesheet.Save();

                var sheets = new Sheets();
                var definedNames = new DefinedNames();
                var usedSheetNames = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);

                for (int s = 0; s < scheduleData.Count; s++)
                {
                    var (sheetName, headers, rows) = scheduleData[s];

                    string safeSheet = string.Join("_",
                        sheetName.Split(new[] { '\\', '/', '?', '*', '[', ']', ':' }));
                    if (safeSheet.Length > 31)
                        safeSheet = safeSheet.Substring(0, 31);

                    // Ensure unique sheet names
                    string uniqueName = safeSheet;
                    int suffix = 2;
                    while (!usedSheetNames.Add(uniqueName))
                    {
                        string suffixStr = $" ({suffix++})";
                        uniqueName = safeSheet.Length + suffixStr.Length > 31
                            ? safeSheet.Substring(0, 31 - suffixStr.Length) + suffixStr
                            : safeSheet + suffixStr;
                    }
                    safeSheet = uniqueName;

                    int totalCols = headers.Count;
                    var rowTypes = ClassifyRows(rows, totalCols);
                    var colTypes = DetectColumnTypes(headers, rows, rowTypes);

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheet = new Worksheet();

                    // SheetProperties: green tab color + fitToPage
                    var sheetProps = new SheetProperties();
                    sheetProps.TabColor = new TabColor { Rgb = "FF008000" };
                    sheetProps.PageSetupProperties = new PageSetupProperties { FitToPage = true };
                    worksheet.Append(sheetProps);

                    // SheetViews: freeze rows 1-2
                    worksheet.Append(new SheetViews(
                        new SheetView(
                            new Pane
                            {
                                VerticalSplit = 2D,
                                TopLeftCell = "A3",
                                ActivePane = PaneValues.BottomLeft,
                                State = PaneStateValues.Frozen
                            },
                            new Selection { Pane = PaneValues.BottomLeft }
                        )
                        {
                            RightToLeft = true,
                            TabSelected = s == 0,
                            WorkbookViewId = 0U
                        }
                    ));

                    // SheetFormatProperties
                    worksheet.Append(new SheetFormatProperties { DefaultRowHeight = 15D });

                    // Column widths
                    var cols = new Columns();
                    for (int c = 1; c <= totalCols; c++)
                    {
                        double maxLen = headers[c - 1].Length;
                        foreach (var row in rows)
                        {
                            if (c - 1 < row.Count)
                                maxLen = Math.Max(maxLen, (row[c - 1] ?? "").Length);
                        }
                        double width = Math.Max(12, Math.Min(40, maxLen * 1.3 + 2));
                        cols.Append(new Column { Min = (uint)c, Max = (uint)c, Width = width, CustomWidth = true });
                    }
                    worksheet.Append(cols);

                    var sheetData = new SheetData();

                    // Row 1: Title (merged across all columns, 3 border styles)
                    var titleRow = new Row { RowIndex = 1U, CustomHeight = true, Height = 30D };
                    for (int c = 1; c <= totalCols; c++)
                    {
                        uint titleStyle;
                        if (totalCols == 1)
                            titleStyle = STL_TITLE_L;
                        else if (c == 1)
                            titleStyle = STL_TITLE_L;
                        else if (c == totalCols)
                            titleStyle = STL_TITLE_R;
                        else
                            titleStyle = STL_TITLE_M;

                        if (c == 1)
                            titleRow.Append(MakeStringCell(1, c, sheetName, titleStyle));
                        else
                            titleRow.Append(new Cell
                            {
                                CellReference = $"{ColLetter(c)}1",
                                StyleIndex = titleStyle
                            });
                    }
                    sheetData.Append(titleRow);

                    // Row 2: Column headers
                    var headerRow = new Row { RowIndex = 2U, CustomHeight = true, Height = 28D };
                    for (int c = 0; c < totalCols; c++)
                        headerRow.Append(MakeStringCell(2, c + 1, headers[c], STL_HEADER));
                    sheetData.Append(headerRow);

                    // Data rows (starting at Excel row 3)
                    int normalIdx = 0;
                    for (int i = 0; i < rows.Count; i++)
                    {
                        int excelRow = i + 3;
                        int rType = rowTypes[i];
                        var dataRow = new Row { RowIndex = (uint)excelRow };

                        if (rType == 0)
                        {
                            dataRow.CustomHeight = true;
                            dataRow.Height = 22D;
                        }

                        for (int c = 0; c < rows[i].Count; c++)
                        {
                            string text = (rows[i][c] ?? "").Trim();
                            bool isAlt = rType == 0 && normalIdx % 2 == 1;
                            ColumnType colType = c < colTypes.Length ? colTypes[c] : ColumnType.Text;

                            CellKind kind;
                            double numVal = 0;

                            if (string.IsNullOrWhiteSpace(text))
                            {
                                kind = CellKind.Empty;
                            }
                            else if (colType == ColumnType.Pct)
                            {
                                bool hadPercentSign = text.EndsWith("%");
                                string numPart = hadPercentSign ? text.TrimEnd('%').Trim() : text;
                                if (double.TryParse(numPart, NumberStyles.Any, CultureInfo.CurrentCulture, out numVal))
                                {
                                    if (numVal == 0)
                                    {
                                        kind = CellKind.Empty;
                                    }
                                    else
                                    {
                                        kind = CellKind.Pct;
                                        if (hadPercentSign && rType == 0) numVal /= 100.0;
                                    }
                                }
                                else
                                {
                                    kind = CellKind.Text;
                                }
                            }
                            else if (colType == ColumnType.Int || colType == ColumnType.Dec)
                            {
                                if (TryParseNumber(text, out numVal))
                                {
                                    if (numVal == 0)
                                        kind = CellKind.Empty;
                                    else if (colType == ColumnType.Dec)
                                        kind = CellKind.Dec;
                                    else
                                        kind = CellKind.Int;
                                }
                                else
                                {
                                    kind = CellKind.Text;
                                }
                            }
                            else
                            {
                                kind = CellKind.Text;
                            }

                            uint style = GetCellStyle(rType, isAlt, kind);

                            if (kind == CellKind.Int || kind == CellKind.Dec || kind == CellKind.DecZero || kind == CellKind.Pct)
                                dataRow.Append(MakeNumberCell(excelRow, c + 1, numVal, style));
                            else
                                dataRow.Append(MakeStringCell(excelRow, c + 1, text, style));
                        }

                        if (rType == 0) normalIdx++;
                        sheetData.Append(dataRow);
                    }

                    worksheet.Append(sheetData);

                    // Merge title row
                    if (totalCols > 1)
                    {
                        worksheet.Append(new MergeCells(
                            new MergeCell { Reference = $"A1:{ColLetter(totalCols)}1" }
                        ));
                    }

                    // Page setup
                    worksheet.Append(new PageMargins
                    {
                        Left = 0.75D, Right = 0.75D, Top = 0.75D,
                        Bottom = 0.5D, Header = 0.5D, Footer = 0.75D
                    });
                    worksheet.Append(new PageSetup
                    {
                        Orientation = OrientationValues.Landscape,
                        FitToWidth = 1U,
                        FitToHeight = 0U
                    });

                    worksheetPart.Worksheet = worksheet;
                    worksheetPart.Worksheet.Save();

                    sheets.Append(new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = (uint)(s + 1),
                        Name = safeSheet
                    });

                    definedNames.Append(
                        new DefinedName($"'{safeSheet}'!$1:$2")
                        {
                            Name = "_xlnm.Print_Titles",
                            LocalSheetId = (uint)s
                        });
                }

                workbookPart.Workbook.Append(sheets);
                workbookPart.Workbook.Append(definedNames);
                workbookPart.Workbook.Save();
            }
        }

        #endregion

        #region Stylesheet

        private Stylesheet BuildStylesheet()
        {
            // 7 Fonts
            var fonts = new Fonts(
                MakeFont("Calibri", 11),                          // 0: Default
                MakeFont("Arial", 14, true, "FFFFFFFF"),          // 1: Title (bold white)
                MakeFont("Arial", 10, true, "FFFFFFFF"),          // 2: Header (bold white)
                MakeFont("Arial", 10),                            // 3: Normal data
                MakeFont("Arial", 10, false, "FFBFBFBF"),         // 4: Gray (zeros)
                MakeFont("Arial", 10, true, "FF1F4E79"),          // 5: Subtotal (bold navy)
                MakeFont("Arial", 11, true, "FF1F4E79")           // 6: Grand total (bold navy)
            ) { Count = 7 };

            // 8 Fills (first two required by spec)
            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                MakeSolidFill("FF1F4E79"),   // 2: Dark navy (title)
                MakeSolidFill("FF2E75B6"),   // 3: Medium blue (header)
                MakeSolidFill("FFFFFFFF"),   // 4: White
                MakeSolidFill("FFF2F7FB"),   // 5: Alt row
                MakeSolidFill("FFDCE6F1"),   // 6: Subtotal
                MakeSolidFill("FFB4C6E7")    // 7: Grand total
            ) { Count = 8 };

            // 7 Borders
            var borders = new Borders(
                EmptyBorder(),                                                                            // 0: None
                MakeBorderFull("FF000000", "FF000000", "FF000000", "FF000000", false, false),              // 1: Title-L (thin black all)
                MakeBorderPartial(null, null, "FF000000", "FF000000"),                                     // 2: Title-M (no L/R, thin T/B)
                MakeBorderPartial(null, "FF000000", "FF000000", "FF000000"),                               // 3: Title-R (no L, thin R/T/B)
                MakeBorderFull("FF1F4E79", "FF1F4E79", "FF1F4E79", "FF1F4E79", false, true),              // 4: Header (thin navy, medium bottom)
                MakeBorderFull("FFB4C6E7", "FFB4C6E7", "FFB4C6E7", "FFB4C6E7", false, false),            // 5: Data (thin light blue)
                MakeBorderFull("FFB4C6E7", "FFB4C6E7", "FF2E75B6", "FF2E75B6", true, true)               // 6: Summary (thin LR, med TB)
            ) { Count = 7 };

            // CellStyleFormats (cellStyleXfs)
            var cellStyleFormats = new CellStyleFormats(
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }
            ) { Count = 1 };

            // 27 CellFormats (cellXfs) — uses built-in numFmtIds: 3=#,##0  4=#,##0.00  9=0%
            var cf = new CellFormats(
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },                       //  0: DEFAULT
                MakeCF(2, 3, 4, 0, HorizontalAlignmentValues.Center, true, true),                             //  1: HEADER
                MakeCF(3, 4, 5, 3, HorizontalAlignmentValues.Center, false, true),                            //  2: W_INT
                MakeCF(4, 4, 5, 4, HorizontalAlignmentValues.Center, false, true),                            //  3: W_DEC_ZERO (gray font)
                MakeCF(3, 4, 5, 4, HorizontalAlignmentValues.Center, false, true),                            //  4: W_DEC
                MakeCF(3, 4, 5, 0, HorizontalAlignmentValues.Right, true, true),                              //  5: W_TEXT
                MakeCF(3, 4, 5, 9, HorizontalAlignmentValues.Center, false, true),                            //  6: W_PCT
                MakeCF(3, 4, 5, 0, HorizontalAlignmentValues.Center, true, true),                             //  7: W_CENTER_WRAP
                MakeCF(3, 5, 5, 3, HorizontalAlignmentValues.Center, false, true),                            //  8: A_INT
                MakeCF(3, 5, 5, 4, HorizontalAlignmentValues.Center, false, true),                            //  9: A_DEC
                MakeCF(3, 5, 5, 0, HorizontalAlignmentValues.Right, true, true),                              // 10: A_TEXT
                MakeCF(3, 5, 5, 9, HorizontalAlignmentValues.Center, false, true),                            // 11: A_PCT
                MakeCF(3, 5, 5, 0, HorizontalAlignmentValues.Center, true, true),                             // 12: A_CENTER_WRAP
                MakeCF(5, 6, 6, 3, HorizontalAlignmentValues.Center, false, true),                            // 13: SUB_INT
                MakeCF(5, 6, 6, 4, HorizontalAlignmentValues.Center, false, true),                            // 14: SUB_DEC
                MakeCF(5, 6, 6, 0, HorizontalAlignmentValues.Right, true, true),                              // 15: SUB_TEXT
                MakeCF(5, 6, 6, 0, HorizontalAlignmentValues.Center, false, true),                            // 16: SUB_CENTER
                MakeCF(5, 6, 6, 0, HorizontalAlignmentValues.Center, true, true),                             // 17: SUB_CENTER_WRAP
                MakeCF(4, 5, 5, 4, HorizontalAlignmentValues.Center, false, true),                            // 18: A_DEC_ZERO (gray font)
                MakeCF(6, 7, 6, 3, HorizontalAlignmentValues.Center, false, true),                            // 19: GT_INT
                MakeCF(6, 7, 6, 4, HorizontalAlignmentValues.Center, false, true),                            // 20: GT_DEC
                MakeCF(6, 7, 6, 0, HorizontalAlignmentValues.Right, true, true),                              // 21: GT_TEXT
                MakeCF(6, 7, 6, 0, HorizontalAlignmentValues.Center, false, true),                            // 22: GT_CENTER
                MakeCF(6, 7, 6, 0, HorizontalAlignmentValues.Center, true, true),                             // 23: GT_CENTER_WRAP
                MakeCF(1, 2, 1, 0, HorizontalAlignmentValues.Center, true, true),                             // 24: TITLE_L
                MakeCF(1, 2, 2, 0, HorizontalAlignmentValues.Center, true, true),                             // 25: TITLE_M
                MakeCF(1, 2, 3, 0, HorizontalAlignmentValues.Center, true, true)                              // 26: TITLE_R
            ) { Count = 27 };

            // CellStyles
            var cellStyles = new CellStyles(
                new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }
            ) { Count = 1 };

            return new Stylesheet(fonts, fills, borders, cellStyleFormats, cf, cellStyles);
        }

        #endregion

        #region Stylesheet Helpers

        private static XlFont MakeFont(string name, double size, bool bold = false, string? colorHex = null)
        {
            var font = new XlFont();
            if (bold) font.Append(new Bold());
            font.Append(new FontSize { Val = size });
            font.Append(new XlColor { Rgb = colorHex ?? "FF000000" });
            font.Append(new FontName { Val = name });
            return font;
        }

        private static Fill MakeSolidFill(string colorHex)
        {
            return new Fill(new PatternFill(
                new ForegroundColor { Rgb = colorHex }) { PatternType = PatternValues.Solid });
        }

        private static XlBorder EmptyBorder()
        {
            return new XlBorder(
                new LeftBorder(), new RightBorder(), new TopBorder(),
                new BottomBorder(), new DiagonalBorder());
        }

        private static XlBorder MakeBorderFull(string left, string right, string top, string bottom,
            bool tMed, bool bMed)
        {
            return new XlBorder(
                new LeftBorder(new XlColor { Rgb = left }) { Style = BorderStyleValues.Thin },
                new RightBorder(new XlColor { Rgb = right }) { Style = BorderStyleValues.Thin },
                new TopBorder(new XlColor { Rgb = top }) { Style = tMed ? BorderStyleValues.Medium : BorderStyleValues.Thin },
                new BottomBorder(new XlColor { Rgb = bottom }) { Style = bMed ? BorderStyleValues.Medium : BorderStyleValues.Thin },
                new DiagonalBorder());
        }

        private static XlBorder MakeBorderPartial(string? left, string? right, string? top, string? bottom)
        {
            return new XlBorder(
                left != null
                    ? new LeftBorder(new XlColor { Rgb = left }) { Style = BorderStyleValues.Thin }
                    : new LeftBorder(),
                right != null
                    ? new RightBorder(new XlColor { Rgb = right }) { Style = BorderStyleValues.Thin }
                    : new RightBorder(),
                top != null
                    ? new TopBorder(new XlColor { Rgb = top }) { Style = BorderStyleValues.Thin }
                    : new TopBorder(),
                bottom != null
                    ? new BottomBorder(new XlColor { Rgb = bottom }) { Style = BorderStyleValues.Thin }
                    : new BottomBorder(),
                new DiagonalBorder());
        }

        private static CellFormat MakeCF(uint fontId, uint fillId, uint borderId, uint numFmtId,
            HorizontalAlignmentValues hAlign, bool wrap, bool rtl)
        {
            var format = new CellFormat
            {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                FormatId = 0,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true,
                ApplyAlignment = true,
                Alignment = new Alignment
                {
                    Horizontal = hAlign,
                    Vertical = VerticalAlignmentValues.Center,
                    WrapText = wrap,
                    ReadingOrder = rtl ? (UInt32Value)2U : null
                }
            };

            if (numFmtId > 0)
            {
                format.NumberFormatId = numFmtId;
                format.ApplyNumberFormat = true;
            }

            return format;
        }

        #endregion

        #region Cell Helpers

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

        private static Cell MakeNumberCell(int row, int col, double value, uint styleIndex)
        {
            return new Cell
            {
                CellReference = $"{ColLetter(col)}{row}",
                DataType = CellValues.Number,
                StyleIndex = styleIndex,
                CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture))
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

        #endregion
    }
}
