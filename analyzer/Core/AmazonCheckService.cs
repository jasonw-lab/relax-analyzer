using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer.Core
{
    /// <summary>
    /// Amazon購入履歴とカード利用明細を照合し、コメント列に商品名を記入するサービス
    /// </summary>
    internal sealed class AmazonCheckService
    {
      // カード利用明細シートの列定義
      private const int CardSheetStartRow = 4;
      private const int CardSheetUseDateColumn = 1;   // A列: 利用日
      private const int CardSheetStoreColumn = 2;     // B列: 利用店名・商品名
        private const int CardSheetAmountColumn = 5;    // E列: 利用金額
        private const int CardSheetCommentColumn = 12;  // L列: コメント

        // amazonシートの列定義
        private const int AmazonSheetStartRow = 2;      // Header row is 1
     private const int AmazonOrderDateColumn = 1;    // A列: Order Date
        private const int AmazonItemShortNameColumn = 3;// C列: Item Short Name
        private const int AmazonAmountColumn = 4;       // D列: 金額

     private readonly IList<string> _warnings;

        public AmazonCheckService(IList<string> warnings)
  {
   _warnings = warnings ?? new List<string>();
        }

        /// <summary>
        /// 指定されたカード利用明細シートに対してAmazon照合を実行
        /// </summary>
        /// <param name="workbook">対象ワークブック</param>
        /// <param name="cardSheetName">カード利用明細シート名（例: "1", "2", ... "12"）</param>
        public void CheckAmazonForSheet(Excel.Workbook workbook, string cardSheetName)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(cardSheetName)) throw new ArgumentException(nameof(cardSheetName));

            // カード利用明細シートを取得
            var cardSheet = SheetHelper.FindSheet(workbook, cardSheetName);
            if (cardSheet == null)
            {
                _warnings.Add($"シート '{cardSheetName}' が見つかりません。");
                return;
            }

            var amazonSheet = SheetHelper.FindSheet(workbook, "amazon");
            if (amazonSheet == null)
            {
                _warnings.Add("'amazon' シートが見つかりません。先にAmazon CSV サマリを作成してください。");
                return;
            }

            var amazonOrders = LoadAmazonOrders(amazonSheet);
            if (amazonOrders.Count == 0)
            {
                _warnings.Add("'amazon' シートにデータがありません。");
                return;
            }

            ProcessCardSheet(cardSheet, amazonOrders);
        }

        /// <summary>
        /// 全てのカード利用明細シート（1〜12）に対してAmazon照合を実行
        /// </summary>
        public void CheckAmazonForAllSheets(Excel.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            var amazonSheet = SheetHelper.FindSheet(workbook, "amazon");
            if (amazonSheet == null)
            {
                _warnings.Add("'amazon' シートが見つかりません。先にAmazon CSV サマリを作成してください。");
                return;
            }

            var amazonOrders = LoadAmazonOrders(amazonSheet);
            if (amazonOrders.Count == 0)
            {
                _warnings.Add("'amazon' シートにデータがありません。");
                return;
            }

            var processedCount = 0;
            for (var month = 1; month <= 12; month++)
            {
                var sheetName = month.ToString(CultureInfo.InvariantCulture);
                var sheet = SheetHelper.FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    continue;
                }

                ProcessCardSheet(sheet, amazonOrders);
                processedCount++;
            }

            if (processedCount == 0)
            {
                _warnings.Add("カード利用明細シート（1〜12）が見つかりませんでした。");
            }
        }

        private List<AmazonOrder> LoadAmazonOrders(Excel.Worksheet amazonSheet)
        {
   var orders = new List<AmazonOrder>();

            var usedRange = amazonSheet.UsedRange;
            if (usedRange == null)
          {
         return orders;
        }

   var lastRow = usedRange.Rows.Count;
 if (lastRow < AmazonSheetStartRow)
  {
          return orders;
            }

            // データ範囲を一括取得（パフォーマンス向上）
    var dateRange = (Excel.Range)amazonSheet.Range[
    amazonSheet.Cells[AmazonSheetStartRow, AmazonOrderDateColumn],
           amazonSheet.Cells[lastRow, AmazonOrderDateColumn]];
   var dateValues = dateRange.Value2 as object[,];

            var itemRange = (Excel.Range)amazonSheet.Range[
                amazonSheet.Cells[AmazonSheetStartRow, AmazonItemShortNameColumn],
        amazonSheet.Cells[lastRow, AmazonItemShortNameColumn]];
            var itemValues = itemRange.Value2 as object[,];

    var amountRange = (Excel.Range)amazonSheet.Range[
     amazonSheet.Cells[AmazonSheetStartRow, AmazonAmountColumn],
             amazonSheet.Cells[lastRow, AmazonAmountColumn]];
  var amountValues = amountRange.Value2 as object[,];

  if (dateValues == null || itemValues == null || amountValues == null)
  {
     return orders;
            }

      var rowCount = dateValues.GetLength(0);

      for (var i = 1; i <= rowCount; i++)
        {
                var dateObj = dateValues[i, 1];
                var itemObj = itemValues[i, 1];
            var amountObj = amountValues[i, 1];

        var itemStr = itemObj == null ? string.Empty : Convert.ToString(itemObj).Trim();

       if (dateObj == null || string.IsNullOrEmpty(itemStr))
       {
  continue;
   }

       var orderDate = ParseDateFromExcel(dateObj);
                if (!orderDate.HasValue)
      {
               continue;
     }

var amount = ParseAmountFromExcel(amountObj);

      orders.Add(new AmazonOrder
            {
           OrderDate = orderDate.Value,
ItemShortName = itemStr,
       Amount = amount
     });
            }

            _warnings.Add($"[デバッグ] amazon シート読み込み: {orders.Count} 件");

            return orders;
        }

        private void ProcessCardSheet(Excel.Worksheet cardSheet, List<AmazonOrder> amazonOrders)
        {
  var usedRange = cardSheet.UsedRange;
        if (usedRange == null)
            {
     return;
         }

      var lastRow = usedRange.Rows.Count;
  if (lastRow < CardSheetStartRow)
     {
     return;
            }

          var totalRows = lastRow - CardSheetStartRow + 1;

 // データ範囲を一括取得
  var useDateRange = (Excel.Range)cardSheet.Range[
    cardSheet.Cells[CardSheetStartRow, CardSheetUseDateColumn],
      cardSheet.Cells[lastRow, CardSheetUseDateColumn]];
   var useDateValues = useDateRange.Value2 as object[,];

   var storeRange = (Excel.Range)cardSheet.Range[
                cardSheet.Cells[CardSheetStartRow, CardSheetStoreColumn],
                cardSheet.Cells[lastRow, CardSheetStoreColumn]];
          var storeValues = storeRange.Value2 as object[,];

  var amountRange = (Excel.Range)cardSheet.Range[
                cardSheet.Cells[CardSheetStartRow, CardSheetAmountColumn],
                cardSheet.Cells[lastRow, CardSheetAmountColumn]];
       var amountValues = amountRange.Value2 as object[,];

       var commentRange = (Excel.Range)cardSheet.Range[
         cardSheet.Cells[CardSheetStartRow, CardSheetCommentColumn],
     cardSheet.Cells[lastRow, CardSheetCommentColumn]];
  var commentValues = commentRange.Value2 as object[,];

         if (useDateValues == null || storeValues == null || amountValues == null)
    {
    return;
            }

          if (commentValues == null)
  {
  commentValues = new object[totalRows, 1];
        }

      var app = cardSheet.Application;
            var prevCalc = app.Calculation;
            var prevScreenUpdating = app.ScreenUpdating;
          var prevEnableEvents = app.EnableEvents;

     var updatedCount = 0;
            var skippedCount = 0;

 try
 {
       if (totalRows > 50)
  {
      app.ScreenUpdating = false;
          app.EnableEvents = false;
     app.Calculation = Excel.XlCalculation.xlCalculationManual;
      }

    for (var i = 1; i <= totalRows; i++)
     {
         var useDateObj = useDateValues[i, 1];
   var storeObj = storeValues[i, 1];
      var cardAmountObj = amountValues[i, 1];
      var commentObj = commentValues[i, 1];

      // L列（コメント欄）が空でない場合はスキップ
      var existingComment = commentObj == null ? string.Empty : Convert.ToString(commentObj).Trim();
      if (!string.IsNullOrEmpty(existingComment))
      {
        skippedCount++;
        continue;
      }

        var storeStr = storeObj == null ? string.Empty : Convert.ToString(storeObj).Trim();

            // AMAZON. が含まれるか確認
     if (!ContainsAmazon(storeStr))
    {
         continue;
           }

       // 利用日をパース
        var useDate = ParseDateFromExcel(useDateObj);
       if (!useDate.HasValue)
   {
       continue;
        }

        // カード利用金額を取得
        var cardAmount = ParseAmountFromExcel(cardAmountObj);

 // 前後1週間以内 かつ 金額が同じAmazon注文を検索
var matchedItems = FindMatchingAmazonOrders(useDate.Value, cardAmount, amazonOrders);
      if (matchedItems.Count > 0)
  {
         // 複数該当する場合は改行区切りで結合
                var comment = string.Join(Environment.NewLine, matchedItems);
   commentValues[i, 1] = comment;
     updatedCount++;
        }
          }

          // 一括書き込み
        commentRange.Value2 = commentValues;

         var statusMessage = $"シート '{cardSheet.Name}': {updatedCount} 件のAmazon商品名を記入しました。";
                if (skippedCount > 0)
                {
                    statusMessage += $" ({skippedCount} 件はコメント欄が既に入力済みのためスキップしました)";
                }
                _warnings.Add(statusMessage);
    }
            finally
       {
       if (totalRows > 50)
     {
      app.Calculation = prevCalc;
     app.EnableEvents = prevEnableEvents;
     app.ScreenUpdating = prevScreenUpdating;
        }
  }
        }

        private static bool ContainsAmazon(string storeStr)
        {
   if (string.IsNullOrEmpty(storeStr))
      {
      return false;
     }

       return storeStr.IndexOf("AMAZON.", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// 前後1週間以内 かつ 金額が同じAmazon注文を検索
        /// </summary>
  private static List<string> FindMatchingAmazonOrders(DateTime useDate, decimal? cardAmount, List<AmazonOrder> amazonOrders)
        {
    var matchedItems = new List<string>();
            var startDate = useDate.AddDays(-7);
    var endDate = useDate.AddDays(7);

     foreach (var order in amazonOrders)
{
             // 日付範囲チェック（前後1週間）
     if (order.OrderDate < startDate || order.OrderDate > endDate)
    {
 continue;
   }

        // 金額チェック（金額が同じデータを抽出）
       if (!IsAmountMatch(cardAmount, order.Amount))
  {
continue;
         }

                matchedItems.Add(order.ItemShortName);
        }

    return matchedItems;
        }

        /// <summary>
  /// 金額が一致するかチェック
    /// </summary>
    private static bool IsAmountMatch(decimal? cardAmount, decimal? amazonAmount)
        {
   // 両方nullの場合は一致とみなさない
 if (!cardAmount.HasValue || !amazonAmount.HasValue)
            {
    return false;
            }

      // 金額が完全一致
            return cardAmount.Value == amazonAmount.Value;
        }

        /// <summary>
        /// Excelセルの値から金額を取得
        /// </summary>
   private static decimal? ParseAmountFromExcel(object cellValue)
        {
      if (cellValue == null)
            {
    return null;
            }

            // 数値の場合
   if (cellValue is double d)
 {
         return (decimal)d;
  }

            // 文字列の場合
   var str = Convert.ToString(cellValue).Trim();
        if (string.IsNullOrEmpty(str))
   {
    return null;
     }

            // カンマ、円記号を除去
   str = str.Replace(",", string.Empty)
        .Replace("¥", string.Empty)
  .Replace("￥", string.Empty)
     .Replace("$", string.Empty);

     if (decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            {
             return result;
      }

   if (decimal.TryParse(str, NumberStyles.Any, CultureInfo.CurrentCulture, out result))
            {
     return result;
     }

         return null;
        }

        private static DateTime? ParseDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr))
     {
        return null;
       }

            // yyyy-MM-dd 形式
        if (DateTime.TryParseExact(dateStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt1))
            {
                return dt1;
 }

    // yyyy/MM/dd 形式（2桁）
            if (DateTime.TryParseExact(dateStr, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt2))
            {
        return dt2;
            }

          // yyyy/M/d 形式（1桁対応: 2025/9/16）
            if (DateTime.TryParseExact(dateStr, "yyyy/M/d", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt5))
        {
  return dt5;
            }

            // その他の形式
    if (DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt3))
    {
    return dt3;
            }

        if (DateTime.TryParse(dateStr, CultureInfo.CurrentCulture, DateTimeStyles.None, out var dt4))
   {
                return dt4;
}

     return null;
   }

 /// <summary>
    /// Excelセルの値から日付を取得（シリアル値・文字列両対応）
     /// </summary>
        private static DateTime? ParseDateFromExcel(object cellValue)
        {
            if (cellValue == null)
        {
     return null;
          }

         // Excelシリアル値（double）の場合
  if (cellValue is double serialDate)
      {
         try
    {
 return DateTime.FromOADate(serialDate);
      }
 catch
         {
  return null;
         }
 }

            // 文字列の場合
            var dateStr = Convert.ToString(cellValue).Trim();
 return ParseDate(dateStr);
        }

        /// <summary>
        /// Amazon注文データ
      /// </summary>
        private sealed class AmazonOrder
        {
     public DateTime OrderDate { get; set; }
   public string ItemShortName { get; set; }
  public decimal? Amount { get; set; }
    }
    }
}
