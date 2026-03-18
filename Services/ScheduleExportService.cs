using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Professional schedule export using GetCellText — what you see is what you get.
    /// No ElementId tracking, no roundtrip — display-only export with styled formatting.
    /// </summary>
    public class ScheduleExportService
    {
        // Layout constants
        private const int TitleRow = 1;
        private const int MetaProjectRow = 2;
        private const int MetaVersionRow = 3;
        private const int MetaDateRow = 4;
        // Row 5 = spacer
        private const int HeaderRow = 6;
        private const int DataStartRow = 7;

        // Colors
        private static readonly XLColor DarkBlue = XLColor.FromHtml("#1E3A5F");
        private static readonly XLColor LightBorder = XLColor.FromHtml("#D0D0D0");
        private static readonly XLColor NumericBlue = XLColor.FromHtml("#1E3A5F");
        private static readonly XLColor ZeroGrey = XLColor.FromHtml("#999999");
        private static readonly XLColor GrandTotalFill = XLColor.FromHtml("#D4EDDA");
        private static readonly XLColor SubtotalFill = XLColor.FromHtml("#FFF3CD");

        // Group header background colors by depth
        private static readonly XLColor[] GroupDepthColors =
        {
            XLColor.FromHtml("#E8EEF4"), // depth 0
            XLColor.FromHtml("#F0F4F8"), // depth 1
            XLColor.FromHtml("#F5F7FA"), // depth 2+
        };

        private static readonly Regex CurrencyRegex = new(@"^[₪$€£]", RegexOptions.Compiled);
        private static readonly Regex NumericRegex = new(@"^-?[\d,]+\.?\d*$", RegexOptions.Compiled);

        /// <summary>
        /// Exports a schedule to a professionally styled Excel file.
        /// </summary>
        public void ExportScheduleToExcel(ViewSchedule schedule, string filePath)
        {
            var doc = schedule.Document;
            var tableData = schedule.GetTableData();

            // Read all sections
            var headers = ReadHeaders(schedule, tableData);
            var bodyRows = ReadBody(schedule, tableData);

            if (headers.Count == 0)
                throw new InvalidOperationException("Schedule has no columns.");

            int colCount = headers.Count;

            // Detect column types from data
            var columnTypes = DetectColumnTypes(bodyRows, colCount);

            // Write Excel
            string safeSheet = SanitizeSheetName(schedule.Name);
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add(safeSheet);

            // --- Title row ---
            WriteTitleRow(ws, schedule.Name, colCount);

            // --- Metadata rows ---
            WriteMetadata(ws, doc, colCount);

            // --- Header row ---
            WriteHeaderRow(ws, headers);

            // --- Data rows ---
            int grandTotalRow = -1;
            var subtotalRows = new List<int>();
            var groupHeaderRows = new List<int>();

            for (int r = 0; r < bodyRows.Count; r++)
            {
                int excelRow = DataStartRow + r;
                var row = bodyRows[r];

                // Classify the row
                var rowType = ClassifyRow(row, colCount);

                for (int c = 0; c < Math.Min(row.Count, colCount); c++)
                {
                    int excelCol = c + 1;
                    string cellText = row[c];

                    if (rowType == RowType.Data && columnTypes[c] != ColumnType.Text && !string.IsNullOrEmpty(cellText))
                    {
                        // Try to write as number for proper formatting
                        string cleaned = CleanNumericText(cellText);
                        if (double.TryParse(cleaned, System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out double numVal))
                        {
                            ws.Cell(excelRow, excelCol).Value = numVal;
                        }
                        else
                        {
                            ws.Cell(excelRow, excelCol).Value = cellText;
                        }
                    }
                    else
                    {
                        ws.Cell(excelRow, excelCol).Value = cellText;
                    }
                }

                // Style by row type
                switch (rowType)
                {
                    case RowType.GrandTotal:
                        grandTotalRow = excelRow;
                        break;
                    case RowType.Subtotal:
                        subtotalRows.Add(excelRow);
                        break;
                    case RowType.GroupHeader:
                        groupHeaderRows.Add(excelRow);
                        break;
                    case RowType.Data:
                        StyleDataRow(ws, excelRow, colCount, columnTypes, row);
                        break;
                }
            }

            // Apply group header styling
            foreach (int row in groupHeaderRows)
                StyleGroupHeaderRow(ws, row, colCount, 0);

            // Apply subtotal styling
            foreach (int row in subtotalRows)
                StyleSubtotalRow(ws, row, colCount);

            // Apply grand total styling
            if (grandTotalRow > 0)
                StyleGrandTotalRow(ws, grandTotalRow, colCount);

            // Apply numeric/currency column formatting to data rows
            ApplyColumnFormats(ws, columnTypes, DataStartRow, DataStartRow + bodyRows.Count - 1,
                grandTotalRow, subtotalRows, groupHeaderRows);

            // Freeze panes (freeze above data, below headers)
            ws.SheetView.FreezeRows(HeaderRow);

            // Auto-fit columns with bounds
            AutoFitColumns(ws, colCount);

            // Print setup
            ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
            ws.PageSetup.FitToPages(1, 0);

            workbook.SaveAs(filePath);

            Log.Information("Exported schedule '{ScheduleName}' to {FilePath}: {RowCount} body rows, {ColCount} columns",
                schedule.Name, filePath, bodyRows.Count, colCount);
        }

        #region Reading Schedule Data

        private List<string> ReadHeaders(ViewSchedule schedule, TableData tableData)
        {
            var headers = new List<string>();
            var headerSection = tableData.GetSectionData(SectionType.Header);
            var bodySection = tableData.GetSectionData(SectionType.Body);
            int colCount = bodySection.NumberOfColumns;

            // Try header section first
            if (headerSection.NumberOfRows > 0)
            {
                for (int col = 0; col < colCount; col++)
                {
                    string text = schedule.GetCellText(SectionType.Header, 0, col);
                    headers.Add(text ?? "");
                }
            }

            // If headers are empty or all blank, fallback to first body row
            if (headers.Count == 0 || headers.All(string.IsNullOrWhiteSpace))
            {
                headers.Clear();
                if (bodySection.NumberOfRows > 0)
                {
                    for (int col = 0; col < colCount; col++)
                    {
                        string text = schedule.GetCellText(SectionType.Body, 0, col);
                        headers.Add(text ?? "");
                    }
                }
            }

            return headers;
        }

        private List<List<string>> ReadBody(ViewSchedule schedule, TableData tableData)
        {
            var rows = new List<List<string>>();
            var bodySection = tableData.GetSectionData(SectionType.Body);
            int rowCount = bodySection.NumberOfRows;
            int colCount = bodySection.NumberOfColumns;

            for (int row = 0; row < rowCount; row++)
            {
                var cells = new List<string>();
                for (int col = 0; col < colCount; col++)
                {
                    string text = schedule.GetCellText(SectionType.Body, row, col);
                    cells.Add(text ?? "");
                }
                rows.Add(cells);
            }

            return rows;
        }

        #endregion

        #region Row Classification

        private enum RowType { Data, GroupHeader, Subtotal, GrandTotal }

        private RowType ClassifyRow(List<string> row, int colCount)
        {
            int filledCells = 0;
            int emptyCells = 0;
            string firstNonEmpty = "";

            for (int c = 0; c < Math.Min(row.Count, colCount); c++)
            {
                if (string.IsNullOrWhiteSpace(row[c]))
                    emptyCells++;
                else
                {
                    filledCells++;
                    if (string.IsNullOrEmpty(firstNonEmpty))
                        firstNonEmpty = row[c];
                }
            }

            // Grand total: contains "Grand Total" or "סה״כ כללי" or "סה"כ"
            if (firstNonEmpty.Contains("Grand Total", StringComparison.OrdinalIgnoreCase) ||
                firstNonEmpty.Contains("סה״כ כללי") || firstNonEmpty.Contains("סה\"כ כללי"))
                return RowType.GrandTotal;

            // Single filled cell in a multi-column schedule = group header or subtotal
            if (colCount > 2 && filledCells <= 2 && emptyCells > filledCells)
            {
                // Subtotal: contains ":" or total keywords
                if (firstNonEmpty.Contains(":") || firstNonEmpty.Contains("סה״כ") ||
                    firstNonEmpty.Contains("סה\"כ") ||
                    firstNonEmpty.Contains("Subtotal", StringComparison.OrdinalIgnoreCase) ||
                    firstNonEmpty.Contains("Total", StringComparison.OrdinalIgnoreCase))
                    return RowType.Subtotal;

                return RowType.GroupHeader;
            }

            return RowType.Data;
        }

        #endregion

        #region Column Type Detection

        private enum ColumnType { Text, Numeric, Currency }

        private List<ColumnType> DetectColumnTypes(List<List<string>> rows, int colCount)
        {
            var types = new List<ColumnType>();
            for (int c = 0; c < colCount; c++)
            {
                int numericCount = 0;
                int currencyCount = 0;
                int textCount = 0;
                int sampleCount = 0;

                foreach (var row in rows)
                {
                    if (c >= row.Count) continue;
                    string val = row[c].Trim();
                    if (string.IsNullOrEmpty(val)) continue;

                    // Skip non-data rows for detection
                    var rowType = ClassifyRow(row, colCount);
                    if (rowType != RowType.Data) continue;

                    sampleCount++;
                    if (CurrencyRegex.IsMatch(val))
                        currencyCount++;
                    else if (NumericRegex.IsMatch(val.Replace(",", "")))
                        numericCount++;
                    else
                        textCount++;

                    if (sampleCount >= 20) break; // sample enough
                }

                if (sampleCount == 0)
                    types.Add(ColumnType.Text);
                else if (currencyCount > sampleCount / 2)
                    types.Add(ColumnType.Currency);
                else if (numericCount > sampleCount / 2)
                    types.Add(ColumnType.Numeric);
                else
                    types.Add(ColumnType.Text);
            }
            return types;
        }

        #endregion

        #region Excel Writing

        private void WriteTitleRow(IXLWorksheet ws, string title, int colCount)
        {
            var cell = ws.Cell(TitleRow, 1);
            cell.Value = title;
            cell.Style.Font.FontSize = 18;
            cell.Style.Font.FontColor = DarkBlue;
            cell.Style.Font.Bold = true;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(TitleRow).Height = 30;

            if (colCount > 1)
                ws.Range(TitleRow, 1, TitleRow, colCount).Merge();
        }

        private void WriteMetadata(IXLWorksheet ws, Document doc, int colCount)
        {
            // Project Name
            WriteMetaRow(ws, MetaProjectRow, "Project:", doc.ProjectInformation?.Name ?? "Unknown", colCount);

            // Version / Number
            string version = doc.ProjectInformation?.Number ?? "";
            if (!string.IsNullOrEmpty(version))
                WriteMetaRow(ws, MetaVersionRow, "Project Number:", version, colCount);

            // Export Date
            WriteMetaRow(ws, MetaDateRow, "Exported:", DateTime.Now.ToString("yyyy-MM-dd HH:mm"), colCount);
        }

        private void WriteMetaRow(IXLWorksheet ws, int row, string label, string value, int colCount)
        {
            ws.Cell(row, 1).Value = label;
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.FontColor = XLColor.DarkGray;

            ws.Cell(row, 2).Value = value;
            ws.Cell(row, 2).Style.Font.FontColor = XLColor.DarkGray;
        }

        private void WriteHeaderRow(IXLWorksheet ws, List<string> headers)
        {
            for (int c = 0; c < headers.Count; c++)
            {
                int col = c + 1;
                var cell = ws.Cell(HeaderRow, col);
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = DarkBlue;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.BottomBorderColor = LightBorder;
                cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.TopBorderColor = LightBorder;
                cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.LeftBorderColor = LightBorder;
                cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                cell.Style.Border.RightBorderColor = LightBorder;
            }
            ws.Row(HeaderRow).Height = 25;
        }

        private void StyleDataRow(IXLWorksheet ws, int excelRow, int colCount,
            List<ColumnType> columnTypes, List<string> rowData)
        {
            for (int c = 0; c < colCount; c++)
            {
                int col = c + 1;
                var cell = ws.Cell(excelRow, col);
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Hair;
                cell.Style.Border.BottomBorderColor = LightBorder;
            }
        }

        private void StyleGroupHeaderRow(IXLWorksheet ws, int excelRow, int colCount, int depth)
        {
            int depthIndex = Math.Min(depth, GroupDepthColors.Length - 1);
            var range = ws.Range(excelRow, 1, excelRow, colCount);
            range.Merge();
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = depth == 0 ? 12 : 11;
            range.Style.Fill.BackgroundColor = GroupDepthColors[depthIndex];
            range.Style.Alignment.Indent = depth + 1;
        }

        private void StyleSubtotalRow(IXLWorksheet ws, int excelRow, int colCount)
        {
            var range = ws.Range(excelRow, 1, excelRow, colCount);
            range.Merge();
            range.Style.Font.Italic = true;
            range.Style.Fill.BackgroundColor = SubtotalFill;
        }

        private void StyleGrandTotalRow(IXLWorksheet ws, int excelRow, int colCount)
        {
            var range = ws.Range(excelRow, 1, excelRow, colCount);
            range.Merge();
            range.Style.Font.Bold = true;
            range.Style.Font.FontSize = 12;
            range.Style.Fill.BackgroundColor = GrandTotalFill;
            range.Style.Border.TopBorder = XLBorderStyleValues.Medium;
            range.Style.Border.TopBorderColor = DarkBlue;
            range.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            range.Style.Border.BottomBorderColor = DarkBlue;
        }

        private void ApplyColumnFormats(IXLWorksheet ws, List<ColumnType> columnTypes,
            int firstDataRow, int lastDataRow, int grandTotalRow, List<int> subtotalRows,
            List<int> groupHeaderRows)
        {
            var specialRows = new HashSet<int>(subtotalRows);
            specialRows.UnionWith(groupHeaderRows);
            if (grandTotalRow > 0) specialRows.Add(grandTotalRow);

            for (int c = 0; c < columnTypes.Count; c++)
            {
                if (columnTypes[c] == ColumnType.Text) continue;

                int col = c + 1;
                for (int row = firstDataRow; row <= lastDataRow; row++)
                {
                    if (specialRows.Contains(row)) continue;

                    var cell = ws.Cell(row, col);
                    if (cell.IsEmpty()) continue;

                    if (columnTypes[c] == ColumnType.Currency)
                    {
                        cell.Style.NumberFormat.Format = "₪#,##0.00";
                        if (cell.TryGetValue(out double val))
                        {
                            cell.Style.Font.FontColor = val == 0 ? ZeroGrey : NumericBlue;
                        }
                    }
                    else // Numeric
                    {
                        cell.Style.NumberFormat.Format = "#,##0.000";
                        cell.Style.Font.FontColor = NumericBlue;
                    }
                }
            }
        }

        #endregion

        #region Helpers

        private string CleanNumericText(string text)
        {
            // Remove currency symbols and thousands separators for parsing
            string cleaned = text.Trim();
            cleaned = CurrencyRegex.Replace(cleaned, "");
            cleaned = cleaned.Replace(",", "").Trim();
            return cleaned;
        }

        private void AutoFitColumns(IXLWorksheet ws, int colCount)
        {
            const double minWidth = 8;
            const double maxWidth = 50;

            for (int c = 1; c <= colCount; c++)
            {
                ws.Column(c).AdjustToContents();
                double width = ws.Column(c).Width;
                if (width < minWidth) ws.Column(c).Width = minWidth;
                else if (width > maxWidth) ws.Column(c).Width = maxWidth;
            }
        }

        private static string SanitizeSheetName(string name)
        {
            string safe = string.Join("_", name.Split(new[] { '\\', '/', '?', '*', '[', ']', ':' }));
            if (safe.Length > 31)
                safe = safe.Substring(0, 31);
            return safe;
        }

        #endregion
    }
}
