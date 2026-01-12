using System;
using System.Collections.Generic;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer.Core
{
    /// <summary>
    /// 年間消費シート更新サービス
    /// 月別シート（1〜12）の消費種類別金額を集計し、年間シートのマトリクス形式で表示する
    /// </summary>
    internal sealed class YearSheetService
    {
        // 月別シートの列定義
        private const int MonthSheetStartRow = 4;
        private const int MonthSheetTypeColumn = 11;    // K列: 消費種類
        private const int MonthSheetAmountColumn = 5;   // E列: 利用金額

        // 年間シートの列定義
        private const int YearSheetStartRow = 3;
        private const int YearSheetTypeColumn = 3;      // C列: 消費種類
        private const int YearSheetFirstMonthColumn = 4;// D列: 1月（D〜O列が1〜12月）

        private readonly IList<string> _warnings;

        public YearSheetService(IList<string> warnings)
        {
            _warnings = warnings ?? new List<string>();
        }

        /// <summary>
        /// 年間シート更新処理のエントリポイント
        /// </summary>
        /// <param name="workbook">対象ワークブック</param>
        /// <returns>更新結果（処理済みシート数、更新済み消費種類数）</returns>
        public YearSheetUpdateResult UpdateYearSheet(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = new YearSheetUpdateResult();

            // 年間シートの存在確認
            var yearSheet = SheetHelper.FindSheet(workbook, "年間");
            if (yearSheet == null)
            {
                throw new InvalidOperationException("「年間」シートが見つかりません。先にシートを作成してください。");
            }

            var app = workbook.Application;
            var prevCalc = app.Calculation;
            var prevScreenUpdating = app.ScreenUpdating;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                // 月別データ集計（1〜12）
                var processedCount = 0;
                var monthlyData = CollectMonthlyData(workbook, ref processedCount);
                result.ProcessedSheetCount = processedCount;

                // 年間シート更新
                result.UpdatedTypeCount = UpdateYearSheetMatrix(yearSheet, monthlyData);
            }
            finally
            {
                app.Calculation = prevCalc;
                app.EnableEvents = prevEnableEvents;
                app.ScreenUpdating = prevScreenUpdating;
            }

            return result;
        }

        /// <summary>
        /// 月別シート（1〜12）からデータを集計
        /// </summary>
        /// <param name="workbook">対象ワークブック</param>
        /// <param name="processedSheetCount">処理済みシート数（ref）</param>
        /// <returns>消費種類ごとの月別金額データ（Dictionary<消費種類, 月別金額配列[0-11]>）</returns>
        private Dictionary<string, decimal[]> CollectMonthlyData(Excel.Workbook workbook, ref int processedSheetCount)
        {
            var data = new Dictionary<string, decimal[]>(StringComparer.OrdinalIgnoreCase);

            for (var month = 1; month <= 12; month++)
            {
                var sheetName = month.ToString(CultureInfo.InvariantCulture);
                var sheet = SheetHelper.FindSheet(workbook, sheetName);

                if (sheet == null)
                {
                    continue;
                }

                try
                {
                    var usedRange = sheet.UsedRange;
                    if (usedRange == null)
                    {
                        continue;
                    }

                    var lastRow = usedRange.Rows.Count;
                    if (lastRow < MonthSheetStartRow)
                    {
                        continue;
                    }

                    // K列（消費種類）とE列（利用金額）を一括取得
                    var kRange = (Excel.Range)sheet.Range[
                        sheet.Cells[MonthSheetStartRow, MonthSheetTypeColumn],
                        sheet.Cells[lastRow, MonthSheetTypeColumn]];
                    var eRange = (Excel.Range)sheet.Range[
                        sheet.Cells[MonthSheetStartRow, MonthSheetAmountColumn],
                        sheet.Cells[lastRow, MonthSheetAmountColumn]];

                    var kValuesObj = kRange.Value2;
                    var eValuesObj = eRange.Value2;

                    if (kValuesObj == null || eValuesObj == null)
                    {
                        continue;
                    }

                    var kValues = kValuesObj as object[,];
                    var eValues = eValuesObj as object[,];

                    if (kValues == null || eValues == null)
                    {
                        continue;
                    }

                    var totalRows = kValues.GetLength(0);

                    // 各行の消費種類と金額を集計
                    for (var i = 1; i <= totalRows; i++)
                    {
                        var typeRaw = kValues[i, 1];
                        if (typeRaw == null)
                        {
                            continue;
                        }

                        var type = Convert.ToString(typeRaw).Trim();
                        if (string.IsNullOrEmpty(type))
                        {
                            continue;
                        }

                        var amountRaw = eValues[i, 1];
                        var amount = ParseAmount(amountRaw);

                        if (!data.ContainsKey(type))
                        {
                            data[type] = new decimal[12];
                        }

                        data[type][month - 1] += amount;
                    }

                    processedSheetCount++;
                }
                catch (Exception ex)
                {
                    _warnings.Add($"シート '{sheetName}': {ex.Message}");
                }
            }

            return data;
        }

        /// <summary>
        /// 年間シートのマトリクスを更新
        /// </summary>
        /// <param name="yearSheet">年間シート</param>
        /// <param name="monthlyData">月別集計データ</param>
        /// <returns>更新済み消費種類数</returns>
        private int UpdateYearSheetMatrix(Excel.Worksheet yearSheet, Dictionary<string, decimal[]> monthlyData)
        {
            var updatedCount = 0;

            try
            {
                var usedRange = yearSheet.UsedRange;
                if (usedRange == null)
                {
                    _warnings.Add("年間シートにデータがありません。");
                    return 0;
                }

                var lastRow = usedRange.Rows.Count;
                if (lastRow < YearSheetStartRow)
                {
                    _warnings.Add("年間シートのデータが不足しています。");
                    return 0;
                }

                // C列（消費種類）を一括取得
                var cRange = (Excel.Range)yearSheet.Range[
                    yearSheet.Cells[YearSheetStartRow, YearSheetTypeColumn],
                    yearSheet.Cells[lastRow, YearSheetTypeColumn]];
                var cValuesObj = cRange.Value2;

                if (cValuesObj == null)
                {
                    _warnings.Add("年間シートのC列（消費種類）が取得できませんでした。");
                    return 0;
                }

                var cValues = cValuesObj as object[,];
                if (cValues == null)
                {
                    _warnings.Add("年間シートのC列（消費種類）の形式が不正です。");
                    return 0;
                }

                var totalRows = cValues.GetLength(0);

                // D〜O列（各月）の既存値を取得（月シートに存在しない消費種類の既存値を保持するため）
                var monthRange = (Excel.Range)yearSheet.Range[
                    yearSheet.Cells[YearSheetStartRow, YearSheetFirstMonthColumn],
                    yearSheet.Cells[lastRow, YearSheetFirstMonthColumn + 11]];
                var monthValuesObj = monthRange.Value2;
                var existingMonthValues = monthValuesObj as object[,];

                // C#配列（0始まり）を作成
                var monthValues = new object[totalRows, 12];

                // 既存値がある場合はコピー（Excel配列は1始まりインデックスに注意）
                if (existingMonthValues != null)
                {
                    var existingRows = existingMonthValues.GetLength(0);
                    var existingCols = existingMonthValues.GetLength(1);
                    var copyRows = Math.Min(existingRows, totalRows);
                    var copyCols = Math.Min(existingCols, 12);
                    
                    for (var r = 1; r <= copyRows; r++)  // Excel配列は1始まり
                    {
                        for (var c = 1; c <= copyCols; c++)  // Excel配列は1始まり
                        {
                            monthValues[r - 1, c - 1] = existingMonthValues[r, c];  // C#配列は0始まり
                        }
                    }
                }

                // 各消費種類に対して月別金額を設定
                for (var i = 1; i <= totalRows; i++)
                {
                    var typeRaw = cValues[i, 1];
                    if (typeRaw == null)
                    {
                        continue;
                    }

                    var type = Convert.ToString(typeRaw).Trim();
                    if (string.IsNullOrEmpty(type))
                    {
                        continue;
                    }

                    // 月シートに存在する消費種類のみ上書き
                    if (monthlyData.ContainsKey(type))
                    {
                        var amounts = monthlyData[type];
                        for (var m = 0; m < 12; m++)
                        {
                            monthValues[i - 1, m] = amounts[m];
                        }
                        updatedCount++;
                    }
                    // 月シートに存在しない消費種類は既存値を保持（monthValues配列をそのまま使用）
                }

                // 一括書き込み
                monthRange.Value2 = monthValues;
            }
            catch (Exception ex)
            {
                _warnings.Add($"年間シート更新エラー: {ex.Message}");
            }

            return updatedCount;
        }

        /// <summary>
        /// Excelセルの値から金額を取得
        /// </summary>
        /// <param name="cellValue">セルの値</param>
        /// <returns>金額（decimal）。解析できない場合は 0</returns>
        private static decimal ParseAmount(object cellValue)
        {
            if (cellValue == null)
            {
                return 0m;
            }

            // 数値の場合
            if (cellValue is double d)
            {
                return (decimal)d;
            }

            // 文字列の場合（カンマ区切り対応）
            var amountStr = Convert.ToString(cellValue).Replace(",", "");
            if (decimal.TryParse(amountStr, out var amount))
            {
                return amount;
            }

            return 0m;
        }
    }

    /// <summary>
    /// 年間シート更新結果
    /// </summary>
    internal sealed class YearSheetUpdateResult
    {
        /// <summary>
        /// 処理済み月シート数
        /// </summary>
        public int ProcessedSheetCount { get; set; }

        /// <summary>
        /// 更新済み消費種類数
        /// </summary>
        public int UpdatedTypeCount { get; set; }
    }
}
