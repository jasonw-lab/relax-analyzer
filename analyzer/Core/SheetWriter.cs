using System;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer.Core
{
    internal sealed class SheetState
    {
        public bool Cleared { get; set; }

        public int RowPointer { get; set; } = 4;
    }

    internal sealed class SheetWriter
    {
        private const int StartRow = 4;
        private const int EndRow = 300;
        private const int ColumnCount = 12;

        private readonly Excel.Application _application;
        private readonly IDictionary<string, SheetState> _sheetStates;
        private readonly IList<string> _warnings;

        public SheetWriter(Excel.Application application, IDictionary<string, SheetState> sheetStates, IList<string> warnings)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _sheetStates = sheetStates ?? throw new ArgumentNullException(nameof(sheetStates));
            _warnings = warnings ?? new List<string>();
        }

        public void ApplyBatches(IEnumerable<SheetImportBatch> batches)
        {
            if (batches == null)
            {
                return;
            }

            foreach (var batch in batches)
            {
                var worksheet = FindWorksheet(batch.SheetName);
                if (worksheet == null)
                {
                    _warnings.Add($"ワークシート '{batch.SheetName}' が見つかりません。");
                    continue;
                }

                var state = GetState(batch.SheetName);
                EnsureCleared(worksheet, state);

                foreach (var chunk in batch.Chunks)
                {
                    WriteChunk(worksheet, state, chunk);
                }
            }
        }

        private Excel.Worksheet FindWorksheet(string sheetName)
        {
            if (_application.Workbooks.Count == 0)
            {
                return null;
            }

            var workbook = _application.ActiveWorkbook ?? _application.Workbooks[1];
            if (workbook == null)
            {
                return null;
            }

            Excel.Worksheet candidate = null;
            try
            {
                candidate = workbook.Worksheets[sheetName];
            }
            catch
            {
                // retry with zero padding (e.g. "01")
                var padded = sheetName?.PadLeft(2, '0');
                if (!string.Equals(padded, sheetName, StringComparison.Ordinal))
                {
                    try
                    {
                        candidate = workbook.Worksheets[padded];
                    }
                    catch
                    {
                        candidate = null;
                    }
                }
            }

            return candidate;
        }

        private SheetState GetState(string sheetName)
        {
            SheetState state;
            if (!_sheetStates.TryGetValue(sheetName, out state))
            {
                state = new SheetState();
                _sheetStates[sheetName] = state;
            }

            if (state.RowPointer < StartRow)
            {
                state.RowPointer = StartRow;
            }

            return state;
        }

        private void EnsureCleared(Excel.Worksheet sheet, SheetState state)
        {
            if (state.Cleared)
            {
                return;
            }

            var range = sheet.Range["A" + StartRow, "L" + EndRow];
            range.ClearContents();
            state.Cleared = true;
            state.RowPointer = StartRow;
        }

        private void WriteChunk(Excel.Worksheet sheet, SheetState state, FileImportChunk chunk)
        {
            var availableRows = EndRow - state.RowPointer + 1;
            if (availableRows <= 0)
            {
                _warnings.Add($"シート '{sheet.Name}' が満杯のため '{chunk.FileName}' は無視されました。");
                return;
            }

            var rowsToWrite = Math.Min(chunk.Rows.Count, availableRows);
            var buffer = new object[rowsToWrite, ColumnCount];
            for (var i = 0; i < rowsToWrite; i++)
            {
                var row = chunk.Rows[i].Values;
                for (var c = 0; c < ColumnCount; c++)
                {
                    buffer[i, c] = c < row.Length ? row[c] : string.Empty;
                }
            }

            var startRow = state.RowPointer;
            var endRow = startRow + rowsToWrite - 1;
            var target = sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[endRow, ColumnCount]];
            target.Value2 = buffer;

            // shade file header rows
            for (var localRow = 0; localRow < rowsToWrite; localRow++)
            {
                if (!chunk.Rows[localRow].IsFileHeader)
                {
                    continue;
                }

                var headerRowIndex = startRow + localRow;
                var headerRange = sheet.Range[sheet.Cells[headerRowIndex, 1], sheet.Cells[headerRowIndex, ColumnCount]];
                headerRange.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#E6F3FF"));
            }

            if (rowsToWrite < chunk.Rows.Count)
            {
                _warnings.Add($"シート '{sheet.Name}' で {chunk.Rows.Count - rowsToWrite} 行を破棄しました ({chunk.FileName})。");
            }

            state.RowPointer = endRow + 1;
        }
    }
}

