using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;

namespace CellMate
{
    internal sealed class RowBlock
    {
        public int Start;
        public int Count;
        public int End => Start + Count - 1;
    }

    public static class Commands
    {
        private static object undoSheet;
        private static List<RowBlock> undoBlocks;
        private static bool busy;
        internal static void ClearUndo() { undoSheet = null; undoBlocks = null; }

        // Used by both commands and tests: cell areas on the same rows duplicate only once.
        internal static List<RowBlock> MergeRows(IEnumerable<RowBlock> input)
        {
            var result = new List<RowBlock>();
            foreach (var block in input.OrderBy(b => b.Start))
            {
                var last = result.LastOrDefault();
                if (last != null && block.Start <= last.End + 1) last.Count = Math.Max(last.End, block.End) - last.Start + 1;
                else result.Add(new RowBlock { Start = block.Start, Count = block.Count });
            }
            return result;
        }

        [ExcelCommand(Name = "CellMate_Duplicate", Description = "Duplicate selected rows with formulas and formatting")]
        public static void Duplicate() => Run(DuplicateCore);

        private static void DuplicateCore()
        {
            dynamic app = AddIn.Application;
            dynamic sheet = app.ActiveSheet;
            dynamic selection = GetRange(app);
            if (sheet == null || selection == null) throw new InvalidOperationException("Select a cell or one or more rows first.");
            if ((int)app.ActiveWindow.SelectedSheets.Count != 1) throw new InvalidOperationException("Ungroup the worksheets first.");
            if ((bool)sheet.ProtectContents) throw new InvalidOperationException("Unprotect the worksheet first.");
            if ((bool)sheet.FilterMode) throw new InvalidOperationException("Clear active filters first to avoid duplicating hidden rows unintentionally.");

            var selected = new List<RowBlock>();
            for (int i = 1; i <= (int)selection.Areas.Count; i++)
            {
                dynamic area = selection.Areas.Item[i];
                selected.Add(new RowBlock { Start = (int)area.Row, Count = (int)area.Rows.Count });
            }
            var blocks = MergeRows(selected);
            var total = blocks.Sum(b => b.Count);
            int maxRows = (int)sheet.Rows.Count;
            if (total > 10000) throw new InvalidOperationException("Select no more than 10,000 rows per operation.");
            if (blocks.Last().End + total > maxRows) throw new InvalidOperationException("There is not enough space at the bottom of the worksheet to duplicate these rows.");
            // Excel refuses to shift nonempty/formatted end rows off the sheet. Check before mutation.
            dynamic used = sheet.UsedRange;
            if ((int)used.Row + (int)used.Rows.Count - 1 + total > maxRows)
                throw new InvalidOperationException("Data or formatting at the bottom of the worksheet prevents inserting these rows.");
            foreach (var block in blocks)
            {
                dynamic rows = sheet.Rows[block.Start + ":" + block.End];
                dynamic affected = app.Intersect(rows, used);
                if (affected != null && !Equals(affected.MergeCells, false))
                    throw new InvalidOperationException("Duplicating rows containing merged cells is not supported.");
                for (int i = 1; i <= (int)sheet.ListObjects.Count; i++)
                {
                    dynamic table = sheet.ListObjects.Item[i];
                    // Inserting next to a table can implicitly expand it as well.
                    int first = (int)table.Range.Row;
                    int last = first + (int)table.Range.Rows.Count - 1;
                    if (block.Start <= last + 1 && block.End + 1 >= first)
                        throw new InvalidOperationException("Duplicating rows inside or next to an Excel Table is not supported yet. Convert the table to a range first.");
                }
            }

            ClearUndo();
            bool screen = app.ScreenUpdating;
            bool events = app.EnableEvents;
            var inserted = new List<RowBlock>();
            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                foreach (var block in blocks.AsEnumerable().Reverse())
                {
                    int start = block.End + 1;
                    var heights = new double[block.Count];
                    var hidden = new bool[block.Count];
                    for (int i = 0; i < block.Count; i++)
                    {
                        heights[i] = (double)sheet.Rows[block.Start + i].RowHeight;
                        hidden[i] = (bool)sheet.Rows[block.Start + i].Hidden;
                    }
                    sheet.Rows[block.Start + ":" + block.End].Copy();
                    sheet.Rows[start + ":" + (start + block.Count - 1)].Insert(-4121);
                    // Lower inserted blocks have moved down with this insertion.
                    foreach (var prior in inserted) prior.Start += block.Count;
                    inserted.Add(new RowBlock { Start = start, Count = block.Count });
                    app.CutCopyMode = false;
                    for (int i = 0; i < block.Count; i++)
                    {
                        sheet.Rows[start + i].RowHeight = heights[i];
                        sheet.Rows[start + i].Hidden = hidden[i];
                    }
                }
                dynamic copies = null;
                foreach (var block in inserted.OrderBy(b => b.Start))
                {
                    dynamic rows = sheet.Rows[block.Start + ":" + block.End];
                    copies = copies == null ? rows : app.Union(copies, rows);
                }
                copies.Select();
                undoSheet = sheet;
                undoBlocks = inserted;
            }
            catch (Exception original)
            {
                try
                {
                    app.CutCopyMode = false;
                    foreach (var block in inserted.OrderByDescending(b => b.Start)) sheet.Rows[block.Start + ":" + block.End].Delete();
                }
                catch (Exception rollback)
                {
                    throw new InvalidOperationException("The operation failed and could not be fully rolled back. Review the workbook before saving.\n" + original.Message + "\n" + rollback.Message);
                }
                throw;
            }
            finally
            {
                app.CutCopyMode = false;
                app.EnableEvents = events;
                app.ScreenUpdating = screen;
            }
            // Must be the last Excel API call in the command.
            app.OnUndo("Undo Duplicate Rows - CellMate", "CellMate_UndoDuplicate");
        }

        [ExcelCommand(Name = "CellMate_UndoDuplicate")]
        public static void UndoDuplicate() => Run(() =>
        {
            if (undoSheet == null || undoBlocks == null) return;
            dynamic app = AddIn.Application;
            dynamic sheet = undoSheet;
            bool events = app.EnableEvents;
            bool screen = app.ScreenUpdating;
            try
            {
                app.EnableEvents = false;
                app.ScreenUpdating = false;
                sheet.Activate();
                foreach (var block in undoBlocks.OrderByDescending(b => b.Start)) sheet.Rows[block.Start + ":" + block.End].Delete();
            }
            finally { ClearUndo(); app.EnableEvents = events; app.ScreenUpdating = screen; }
        });

        [ExcelCommand(Name = "CellMate_PasteValues", Description = "Paste copied values only")]
        public static void PasteValues() => Run(() =>
        {
            dynamic app = AddIn.Application;
            dynamic range = GetRange(app);
            if (range == null) throw new InvalidOperationException("Select a destination range first.");
            if ((int)app.ActiveWindow.SelectedSheets.Count != 1 || (bool)app.ActiveSheet.ProtectContents)
                throw new InvalidOperationException("Ungroup and unprotect the worksheet first.");
            if ((int)app.CutCopyMode != 1) throw new InvalidOperationException("Copy the source cells in Excel with Ctrl+C first.");
            ClearUndo();
            range.PasteSpecial(-4163); // xlPasteValues, preserving destination formatting.
        });

        [ExcelCommand(Name = "CellMate_AutoFit", Description = "AutoFit selected columns")]
        public static void AutoFit() => Run(() =>
        {
            dynamic app = AddIn.Application;
            dynamic range = GetRange(app);
            if (range == null) throw new InvalidOperationException("Select a range first.");
            if ((int)app.ActiveWindow.SelectedSheets.Count != 1 || (bool)app.ActiveSheet.ProtectContents)
                throw new InvalidOperationException("Ungroup and unprotect the worksheet first.");
            ClearUndo();
            range.EntireColumn.AutoFit();
        });

        [ExcelCommand(Name = "CellMate_Settings", Description = "Customize CellMate shortcuts")]
        public static void Settings() => Run(() =>
        {
            using (var form = new SettingsForm()) form.ShowDialog(new ExcelWindow());
        });

        private static object GetRange(dynamic app)
        {
            if (app.ActiveWorkbook == null) return null;
            dynamic selection = app.Selection;
            try { var count = selection.Cells.CountLarge; return selection; }
            catch { return null; }
        }

        private static void Run(Action action)
        {
            if (busy) return;
            busy = true;
            try { action(); }
            catch (Exception ex) { AddIn.ShowError(ex.Message); }
            finally { busy = false; }
        }
    }
}
