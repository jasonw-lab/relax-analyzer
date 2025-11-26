using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer.Core
{
    internal static class TypeMappingProvider
    {
        public static IReadOnlyList<TypeKeyword> LoadMappings(RelaxAnalyzerConfig config, Excel.Workbook workbook, IList<string> warnings)
        {
            var pairs = LoadFromWorksheet(workbook);
            if (pairs.Count > 0)
            {
                return pairs;
            }

            var csvPath = config?.TypeCsvPath;
            if (string.IsNullOrWhiteSpace(csvPath) || !File.Exists(csvPath))
            {
                warnings?.Add("type.csv が見つかりません。type 判定は実行されません。");
                return Array.Empty<TypeKeyword>();
            }

            return LoadFromCsv(csvPath, warnings);
        }

        private static IReadOnlyList<TypeKeyword> LoadFromWorksheet(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return Array.Empty<TypeKeyword>();
            }

            Excel.Worksheet target = null;
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, "type", StringComparison.OrdinalIgnoreCase))
                {
                    target = sheet;
                    break;
                }
            }

            if (target == null)
            {
                return Array.Empty<TypeKeyword>();
            }

            var items = new List<TypeKeyword>();
            var usedRange = target.UsedRange;
            var rowsRange = usedRange != null ? usedRange.Rows : null;
            var lastRow = rowsRange != null ? rowsRange.Count : 0;
            if (lastRow < 2)
            {
                return Array.Empty<TypeKeyword>();
            }

            for (var row = 2; row <= lastRow; row++)
            {
                var keyword = Convert.ToString(((Excel.Range)target.Cells[row, 1]).Value2)?.Trim();
                var type = Convert.ToString(((Excel.Range)target.Cells[row, 2]).Value2)?.Trim();
                if (string.IsNullOrEmpty(keyword) || string.IsNullOrEmpty(type))
                {
                    continue;
                }

                items.Add(new TypeKeyword(keyword, type));
            }

            return items;
        }

        private static IReadOnlyList<TypeKeyword> LoadFromCsv(string csvPath, IList<string> warnings)
        {
            var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                BadDataFound = null,
                DetectDelimiter = false,
                Delimiter = ","
            };

            using (var reader = new StreamReader(csvPath))
            using (var csv = new CsvReader(reader, cfg))
            {
                var results = new List<TypeKeyword>();
                while (csv.Read())
                {
                    var keyword = csv.GetField(0)?.Trim();
                    var type = csv.GetField(1)?.Trim();
                    if (string.IsNullOrEmpty(keyword) || string.IsNullOrEmpty(type))
                    {
                        continue;
                    }

                    results.Add(new TypeKeyword(keyword, type));
                }

                if (results.Count == 0)
                {
                    warnings?.Add("type.csv に有効なデータが存在しません。");
                }

                return results;
            }
        }
    }
}

