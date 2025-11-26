using System;
using System.IO;
using System.Text.RegularExpressions;

namespace analyzer.Core
{
    internal static class MonthExtractor
    {
        // 対応パターン: enaviYYYYMM(XXXX).csv のみ
        // 例: enavi202510(3034).csv → YYYY=2025, MM=10
        private static readonly Regex Pattern = new Regex(
            @"^enavi(?<year>\d{4})(?<month>\d{2})\(\d+\)\.csv$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static bool TryGetMonthSheet(string filePath, out string sheetName, out string message)
        {
            sheetName = null;
            message = null;
            var fileName = Path.GetFileName(filePath);
            if (string.IsNullOrEmpty(fileName))
            {
                message = $"ファイル名を解析できません: {filePath}";
                return false;
            }

            var match = Pattern.Match(fileName);
            if (!match.Success)
            {
                message = $"ファイル名から月を抽出できません: {fileName}";
                return false;
            }

            var monthToken = match.Groups["month"].Value;
            if (!int.TryParse(monthToken, out var month) || month < 1 || month > 12)
            {
                message = $"抽出した月番号が不正です ({monthToken}) : {fileName}";
                return false;
            }

            sheetName = month.ToString();
            return true;
        }
    }
}

