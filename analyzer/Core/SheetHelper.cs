using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer.Core
{
    /// <summary>
    /// Excelシート操作の共通ヘルパークラス
    /// </summary>
    internal static class SheetHelper
    {
        /// <summary>
        /// ワークブックから指定された名前のシートを検索（大文字小文字無視）
        /// </summary>
        /// <param name="workbook">対象ワークブック</param>
        /// <param name="sheetName">シート名</param>
        /// <returns>見つかったシート。存在しない場合は null</returns>
        public static Excel.Worksheet FindSheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return null;
            }

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }

            return null;
        }
    }
}
