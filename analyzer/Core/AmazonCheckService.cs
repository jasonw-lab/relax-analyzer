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
        private const int CardSheetStartRow = 4;
        private const int CardSheetUseDateColumn = 1;   // A列: 利用日
        private const int CardSheetStoreColumn = 2;     // B列: 利用店名・商品名
        private const int CardSheetCommentColumn = 12;  // L列: コメント

        private const int AmazonSheetStartRow = 2;      // Header row is 1
   private const int AmazonOrderDateColumn = 1;    // A列: Order Date
        private const int AmazonItemShortNameColumn = 3;// C列: Item Short Name

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
            var cardSheet = FindSheet(workbook, cardSheetName);
            if (cardSheet == null)
  {
     _warnings.Add($"シート '{cardSheetName}' が見つかりません。");
         return;
       }

          // amazonシートを取得
 var amazonSheet = FindSheet(workbook, "amazon");
if (amazonSheet == null)
            {
      _warnings.Add("'amazon' シートが見つかりません。先にAmazon CSV サマリを作成してください。");
      return;
            }

            // Amazonデータを読み込み
            var amazonOrders = LoadAmazonOrders(amazonSheet);
            if (amazonOrders.Count == 0)
       {
                _warnings.Add("'amazon' シートにデータがありません。");
       return;
       }

 // カード明細を処理
            ProcessCardSheet(cardSheet, amazonOrders);
        }

     /// <summary>
        /// 全てのカード利用明細シート（1〜12）に対してAmazon照合を実行
      /// </summary>
  public void CheckAmazonForAllSheets(Excel.Workbook workbook)
        {
    if (workbook == null) throw new ArgumentNullException(nameof(workbook));

 // amazonシートを取得
      var amazonSheet = FindSheet(workbook, "amazon");
            if (amazonSheet == null)
       {
   _warnings.Add("'amazon' シートが見つかりません。先にAmazon CSV サマリを作成してください。");
         return;
     }

    // Amazonデータを読み込み
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
   var sheet = FindSheet(workbook, sheetName);
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

        private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string sheetName)
        {
   foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
         {
  return ws;
            }
      }

      return null;
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

            if (dateValues == null || itemValues == null)
     {
           return orders;
      }

            var rowCount = dateValues.GetLength(0);
            for (var i = 1; i <= rowCount; i++)
            {
      var dateObj = dateValues[i, 1];
      var itemObj = itemValues[i, 1];

     var dateStr = dateObj == null ? string.Empty : Convert.ToString(dateObj).Trim();
        var itemStr = itemObj == null ? string.Empty : Convert.ToString(itemObj).Trim();

       if (string.IsNullOrEmpty(dateStr) || string.IsNullOrEmpty(itemStr))
     {
   continue;
}

           var orderDate = ParseDate(dateStr);
      if (!orderDate.HasValue)
      {
        continue;
    }

         orders.Add(new AmazonOrder
             {
               OrderDate = orderDate.Value,
      ItemShortName = itemStr
                });
            }

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

var commentRange = (Excel.Range)cardSheet.Range[
 cardSheet.Cells[CardSheetStartRow, CardSheetCommentColumn],
         cardSheet.Cells[lastRow, CardSheetCommentColumn]];
var commentValues = commentRange.Value2 as object[,];

            if (useDateValues == null || storeValues == null)
   {
                return;
     }

  // コメント配列の初期化
     if (commentValues == null)
          {
          commentValues = new object[totalRows, 1];
  }

  var app = cardSheet.Application;
     var prevCalc = app.Calculation;
      var prevScreenUpdating = app.ScreenUpdating;
   var prevEnableEvents = app.EnableEvents;

 var updatedCount = 0;

  try
            {
                // パフォーマンス最適化
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

         var useDateStr = useDateObj == null ? string.Empty : Convert.ToString(useDateObj).Trim();
        var storeStr = storeObj == null ? string.Empty : Convert.ToString(storeObj).Trim();

           // AMAZON. が含まれるか確認
                if (!ContainsAmazon(storeStr))
    {
      continue;
               }

              // 利用日をパース
           var useDate = ParseDate(useDateStr);
     if (!useDate.HasValue)
      {
       continue;
  }

          // 前後1週間以内のAmazon注文を検索
        var matchedItems = FindMatchingAmazonOrders(useDate.Value, amazonOrders);
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

                _warnings.Add($"シート '{cardSheet.Name}': {updatedCount} 件のAmazon商品名を記入しました。");
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

        private static List<string> FindMatchingAmazonOrders(DateTime useDate, List<AmazonOrder> amazonOrders)
  {
     var matchedItems = new List<string>();
         var startDate = useDate.AddDays(-7);
            var endDate = useDate.AddDays(7);

  foreach (var order in amazonOrders)
          {
      if (order.OrderDate >= startDate && order.OrderDate <= endDate)
    {
         matchedItems.Add(order.ItemShortName);
       }
   }

            return matchedItems;
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

       // yyyy/MM/dd 形式
       if (DateTime.TryParseExact(dateStr, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt2))
         {
     return dt2;
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

      private sealed class AmazonOrder
        {
    public DateTime OrderDate { get; set; }
        public string ItemShortName { get; set; }
     }
    }
}
