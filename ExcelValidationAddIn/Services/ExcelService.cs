using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelValidationAddIn.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidationAddIn.Services
{
    public class SheetSnapshot
    {
        public string WorksheetName { get; set; }
        public List<string> Headers { get; set; } = new List<string>();
        public List<DocumentRow> Rows { get; set; } = new List<DocumentRow>();
    }

    /// <summary>
    /// All direct COM interaction with the active workbook lives here so the
    /// rest of the add-in never touches Microsoft.Office.Interop.Excel
    /// directly. Every Range/Worksheet pulled from COM is released explicitly
    /// - Excel add-ins that skip this reliably leak EXCEL.EXE processes.
    /// </summary>
    public static class ExcelService
    {
        // Cell addresses (no sheet prefix, e.g. "A2:D2") this add-in marked on
        // the previous run, so the next run clears exactly those and nothing
        // the user highlighted or commented themselves.
        private static readonly List<string> PreviousMarkedRanges = new List<string>();
        private static readonly List<string> PreviousCommentCells = new List<string>();

        private static readonly int ErrorFillOle = System.Drawing.ColorTranslator.ToOle(
            System.Drawing.ColorTranslator.FromHtml("#FFC7CE"));
        private static readonly int WarningFillOle = System.Drawing.ColorTranslator.ToOle(
            System.Drawing.ColorTranslator.FromHtml("#FFEB9C"));

        public static SheetSnapshot ReadActiveSheet(Excel.Application app)
        {
            var snapshot = new SheetSnapshot();
            var sheet = app.ActiveSheet as Excel.Worksheet;
            if (sheet == null)
            {
                return snapshot;
            }

            try
            {
                snapshot.WorksheetName = sheet.Name;

                Excel.Range usedRange = sheet.UsedRange;
                try
                {
                    int rowCount = usedRange.Rows.Count;
                    int colCount = usedRange.Columns.Count;
                    if (rowCount < 2)
                    {
                        return snapshot;
                    }

                    int firstRow = usedRange.Row; // 1-based worksheet row of the used range
                    object[,] values = usedRange.Value2 as object[,];
                    if (values == null)
                    {
                        return snapshot;
                    }

                    var headers = new List<string>();
                    for (int c = 1; c <= colCount; c++)
                    {
                        object cellVal = values[1, c];
                        headers.Add(cellVal == null ? string.Empty : cellVal.ToString().Trim());
                    }
                    snapshot.Headers = headers;

                    for (int r = 2; r <= rowCount; r++)
                    {
                        bool rowIsBlank = true;
                        var rowValues = new Dictionary<string, object>();
                        for (int c = 1; c <= colCount; c++)
                        {
                            object cellVal = values[r, c];
                            if (cellVal != null && cellVal.ToString().Length > 0)
                            {
                                rowIsBlank = false;
                            }
                            string header = headers[c - 1];
                            if (!string.IsNullOrEmpty(header))
                            {
                                rowValues[header] = cellVal;
                            }
                        }
                        if (rowIsBlank)
                        {
                            continue;
                        }

                        snapshot.Rows.Add(new DocumentRow
                        {
                            SheetRowIndex = firstRow + r - 1,
                            Values = rowValues
                        });
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(usedRange);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
            }

            return snapshot;
        }

        public static void RenderValidationResults(Excel.Application app, string worksheetName, int columnCount, List<ValidationError> errors)
        {
            Excel.Worksheet sheet = FindWorksheet(app, worksheetName);
            if (sheet == null)
            {
                return;
            }

            try
            {
                foreach (string address in PreviousMarkedRanges)
                {
                    Excel.Range range = sheet.Range[address];
                    try
                    {
                        range.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(range);
                    }
                }
                foreach (string address in PreviousCommentCells)
                {
                    Excel.Range cell = sheet.Range[address];
                    try
                    {
                        cell.ClearComments();
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cell);
                    }
                }
                PreviousMarkedRanges.Clear();
                PreviousCommentCells.Clear();

                var byRow = errors.GroupBy(e => e.SheetRowIndex);
                foreach (var group in byRow)
                {
                    int rowIndex = group.Key;
                    bool warningOnly = group.All(e => e.Severity == ValidationSeverity.Warning);

                    Excel.Range rowRange = sheet.Range[
                        sheet.Cells[rowIndex, 1],
                        sheet.Cells[rowIndex, Math.Max(columnCount, 1)]];
                    try
                    {
                        rowRange.Interior.Color = warningOnly ? WarningFillOle : ErrorFillOle;
                        PreviousMarkedRanges.Add((string)rowRange.get_Address(false, false));
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(rowRange);
                    }

                    Excel.Range commentCell = (Excel.Range)sheet.Cells[rowIndex, 1];
                    try
                    {
                        string text = string.Join("\n", group.Select(e =>
                            (string.IsNullOrEmpty(e.Column) ? "" : e.Column + ": ") + e.Message));
                        commentCell.ClearComments();
                        commentCell.AddComment(text);
                        PreviousCommentCells.Add((string)commentCell.get_Address(false, false));
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(commentCell);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
            }
        }

        /// <summary>Selects and scrolls to the given worksheet row so the user can jump to an error from the task pane.</summary>
        public static void GoToRow(Excel.Application app, string worksheetName, int sheetRowIndex)
        {
            Excel.Worksheet sheet = FindWorksheet(app, worksheetName);
            if (sheet == null)
            {
                return;
            }

            try
            {
                Excel.Range cell = (Excel.Range)sheet.Cells[sheetRowIndex, 1];
                try
                {
                    app.Goto(cell, true);
                }
                finally
                {
                    Marshal.ReleaseComObject(cell);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
            }
        }

        private static Excel.Worksheet FindWorksheet(Excel.Application app, string name)
        {
            Excel.Sheets sheets = app.ActiveWorkbook.Worksheets;
            try
            {
                foreach (Excel.Worksheet ws in sheets)
                {
                    if (ws.Name == name)
                    {
                        return ws;
                    }
                    Marshal.ReleaseComObject(ws);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
            }
            return null;
        }
    }
}
