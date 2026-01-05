using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace analyzer.Core
{
    internal sealed class AmazonOrderSummaryService
    {
        internal sealed class AmazonOrderSummaryResult
        {
            public string OrderHistoryPath { get; set; }
            public string TransactionalInvoicingPath { get; set; }
            public bool TransactionalInvoicingJoined { get; set; }
            public IReadOnlyList<string> Logs { get; set; }
        }

        private static readonly string[] OutputColumns =
        {
            "Order Date",
            "Order ID",
            "Item Short Name",
            "金額",
            "Order Status",
            "Item Name",
            "Short Name",
            "Quantity",
        };

        private static readonly Dictionary<string, string[]> OrderHistoryCandidates = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["Order ID"] = new[] { "orderId", "order_id", "amazonOrderId", "Order ID", "orderid" },
            ["Order Date"] = new[] { "orderDate", "order_date", "purchaseDate", "Order Date", "purchase_date" },
            ["Order Status"] = new[] { "order_status", "orderStatus", "status", "Order Status" },
            ["Item Name"] = new[] { "itemTitle", "title", "productName", "item_name", "Product Name", "Product name", "Productname" },
            ["Quantity"] = new[] { "quantity", "qty", "itemQuantity", "Quantity" },
        };

        private static readonly string[] OrderHistoryAmountCandidates =
        {
            "Total Owed",
            "totalOwed",
            "total_owed",
            "Paid Amount",
            "paidAmount",
            "Grand Total",
            "grandTotal",
            "Total",
            "total",
        };

        private static readonly string[] TransactionalInvoicingAmountCandidates =
        {
            "totalAmount",
            "paidAmount",
            "grandTotal",
            "total",
            "Total Owed",
            "total_owed",
        };

        private static readonly string[] CancelKeywords = { "cancel", "cancell", "キャンセル", "取消" };

        private static readonly Regex[] BracketPatterns =
        {
            new Regex(@"\([^)]*\)", RegexOptions.Compiled),
            new Regex(@"\[[^\]]*\]", RegexOptions.Compiled),
            new Regex(@"（[^）]*）", RegexOptions.Compiled),
            new Regex(@"【[^】]*】", RegexOptions.Compiled),
        };

        private static readonly Regex CapacityPattern = new Regex(@"\b\d+(?:\.\d+)?\s*(ml|mL|l|L|g|kg|GB|TB|枚|本|個|袋|箱)\b", RegexOptions.Compiled);

        public AmazonOrderSummaryResult Generate(string inputPathOrDirectory, string outputCsvPath)
        {
            if (string.IsNullOrWhiteSpace(inputPathOrDirectory)) throw new ArgumentException(nameof(inputPathOrDirectory));
            if (string.IsNullOrWhiteSpace(outputCsvPath)) throw new ArgumentException(nameof(outputCsvPath));

            var logs = new List<string>();

            var orderHistoryPath = FindLatestFile(inputPathOrDirectory, "Retail.OrderHistory*.csv", required: true);
            var transactionalPath = FindLatestFile(GetSiblingDirectory(orderHistoryPath), "Retail.TransactionalInvoicing*.csv", required: false);

            logs.Add($"OrderHistory: {orderHistoryPath}");
            logs.Add($"TransactionalInvoicing: {(transactionalPath ?? "(not found)")}");

            var orderHistory = ReadCsv(orderHistoryPath);
            logs.Add($"OrderHistory rows: {orderHistory.Rows.Count}");

            var mapping = InferMapping(orderHistory.Headers);
            foreach (var kv in mapping)
            {
                logs.Add($"Mapped(OrderHistory) {kv.Key} => {kv.Value}");
            }

            // amount column (OrderHistory) is optional
            string ohAmountColumn = FindFirstHeader(orderHistory.Headers, OrderHistoryAmountCandidates);

            var amountByOrderIdFromOH = string.IsNullOrEmpty(ohAmountColumn)
                ? new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
                : BuildAmountByOrderId(orderHistory.Rows, mapping["Order ID"], ohAmountColumn);

            var amountByOrderIdFromTI = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var joined = false;
            if (!string.IsNullOrEmpty(transactionalPath) && File.Exists(transactionalPath))
            {
                var ti = ReadCsv(transactionalPath);

                var tiOrderIdCol = FindFirstHeader(ti.Headers, OrderHistoryCandidates["Order ID"]);
                var tiAmountCol = FindFirstHeader(ti.Headers, TransactionalInvoicingAmountCandidates);

                if (!string.IsNullOrEmpty(tiOrderIdCol) && !string.IsNullOrEmpty(tiAmountCol))
                {
                    amountByOrderIdFromTI = BuildAmountByOrderId(ti.Rows, tiOrderIdCol, tiAmountCol);
                    joined = amountByOrderIdFromTI.Count > 0;
                }
            }

            logs.Add($"TransactionalInvoicing joined: {(joined ? "YES" : "NO")}");

            // Build output rows
            var outputRows = new List<Dictionary<string, string>>();
            foreach (var row in orderHistory.Rows)
            {
                var status = Get(row, mapping["Order Status"]);
                if (IsCancelStatus(status))
                {
                    continue;
                }

                var orderId = Get(row, mapping["Order ID"]);
                if (string.IsNullOrWhiteSpace(orderId))
                {
                    continue;
                }

                var itemName = Get(row, mapping["Item Name"]);
                var qty = Get(row, mapping["Quantity"]);
                var date = NormalizeDate(Get(row, mapping["Order Date"]));

                var itemShort = MakeItemShortName(itemName, 30);
                var shortName = MakeShortName(itemShort, 10);

                decimal? amount = null;
                if (amountByOrderIdFromOH.TryGetValue(orderId, out var ohAmount))
                {
                    amount = ohAmount;
                }
                else if (amountByOrderIdFromTI.TryGetValue(orderId, out var tiAmount))
                {
                    amount = tiAmount;
                }

                outputRows.Add(new Dictionary<string, string>
                {
                    ["Order Date"] = date,
                    ["Order ID"] = orderId,
                    ["Item Short Name"] = itemShort,
                    ["金額"] = amount.HasValue ? amount.Value.ToString(CultureInfo.InvariantCulture) : string.Empty,
                    ["Order Status"] = status,
                    ["Item Name"] = itemName,
                    ["Short Name"] = shortName,
                    ["Quantity"] = qty,
                });
            }

            // logs head(5)
            logs.Add("HEAD(5):");
            for (var i = 0; i < Math.Min(5, outputRows.Count); i++)
            {
                var r = outputRows[i];
                logs.Add($" {i + 1}: {string.Join(", ", OutputColumns.Select(c => c + "=" + (r.TryGetValue(c, out var v) ? v : "")))}");
            }

            WriteCsv(outputCsvPath, outputRows);
            logs.Add($"Wrote: {outputCsvPath} (rows={outputRows.Count})");

            return new AmazonOrderSummaryResult
            {
                OrderHistoryPath = orderHistoryPath,
                TransactionalInvoicingPath = transactionalPath,
                TransactionalInvoicingJoined = joined,
                Logs = logs,
            };
        }

        internal static object[,] ReadOutputCsvForPaste(string csvPath)
        {
            if (string.IsNullOrWhiteSpace(csvPath)) throw new ArgumentException(nameof(csvPath));

            var table = ReadCsv(csvPath);
            if (table.Headers.Count == 0)
            {
                return new object[0, 0];
            }

            var rows = table.Rows;

            var outCols = OutputColumns;
            var colCount = outCols.Length;
            var rowCount = rows.Count + 1; // header row
            var data = new object[rowCount, colCount];

            for (var c = 0; c < colCount; c++)
            {
                data[0, c] = outCols[c];
            }

            for (var r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (var c = 0; c < colCount; c++)
                {
                    var key = outCols[c];
                    data[r + 1, c] = row.TryGetValue(key, out var v) ? v : string.Empty;
                }
            }

            return data;
        }

        private static string GetSiblingDirectory(string filePath)
        {
            return Path.GetDirectoryName(filePath) ?? ".";
        }

        private static string FindLatestFile(string pathOrDirectory, string pattern, bool required)
        {
            if (Directory.Exists(pathOrDirectory))
            {
                var files = Directory.GetFiles(pathOrDirectory, pattern).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
                if (files.Length == 0)
                {
                    if (required) throw new FileNotFoundException($"{pattern} not found in directory: {pathOrDirectory}");
                    return null;
                }

                return files[files.Length - 1];
            }

            if (!File.Exists(pathOrDirectory))
            {
                if (required) throw new FileNotFoundException(pathOrDirectory);
                return null;
            }

            return pathOrDirectory;
        }

        private static string InferRequiredHeader(Dictionary<string, int> headerIndex, string logicalName)
        {
            if (!OrderHistoryCandidates.TryGetValue(logicalName, out var candidates))
            {
                throw new InvalidOperationException("Unknown logical header: " + logicalName);
            }

            var found = FindFirstHeader(headerIndex.Keys, candidates);
            if (string.IsNullOrEmpty(found))
            {
                throw new InvalidOperationException($"Required column not found: {logicalName}. Candidates: {string.Join(", ", candidates)}");
            }

            return found;
        }

        private static Dictionary<string, string> InferMapping(IReadOnlyList<string> headers)
        {
            var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in headers)
            {
                if (!headerIndex.ContainsKey(h))
                {
                    headerIndex[h] = headerIndex.Count;
                }
            }

            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var logical in OrderHistoryCandidates.Keys)
            {
                mapping[logical] = InferRequiredHeader(headerIndex, logical);
            }

            return mapping;
        }

        private static string FindFirstHeader(IEnumerable<string> headers, IEnumerable<string> candidates)
        {
            var normalized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var h in headers)
            {
                var key = NormalizeHeader(h);
                if (!normalized.ContainsKey(key))
                {
                    normalized[key] = h;
                }
            }

            foreach (var c in candidates)
            {
                var key = NormalizeHeader(c);
                if (normalized.TryGetValue(key, out var actual))
                {
                    return actual;
                }
            }

            return null;
        }

        private static string NormalizeHeader(string header)
        {
            return Regex.Replace((header ?? string.Empty).Trim(), @"\s+", string.Empty).ToLowerInvariant();
        }

        private static bool IsCancelStatus(string status)
        {
            var s = (status ?? string.Empty).ToLowerInvariant();
            foreach (var k in CancelKeywords)
            {
                if (s.Contains(k.ToLowerInvariant()))
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeDate(string input)
        {
            var raw = (input ?? string.Empty).Trim();
            if (raw.Length == 0) return raw;

            // ISO8601 like2025-11-17T03:12:30.575Z
            if (raw.EndsWith("Z", StringComparison.OrdinalIgnoreCase))
            {
                if (DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var dto))
                {
                    return dto.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                }
            }

            // fallback formats
            if (DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            if (DateTime.TryParse(raw, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
            {
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return raw;
        }

        private static string MakeItemShortName(string itemName, int maxLen)
        {
            var s = (itemName ?? string.Empty).Trim();
            foreach (var pat in BracketPatterns)
            {
                s = pat.Replace(s, string.Empty);
            }

            s = CapacityPattern.Replace(s, string.Empty);
            s = Regex.Replace(s, @"\s+", " ").Trim();

            if (s.Length > maxLen)
            {
                s = s.Substring(0, maxLen).TrimEnd();
            }

            return s;
        }

        private static string MakeShortName(string itemShortName, int maxLen)
        {
            var s = (itemShortName ?? string.Empty).Trim();
            if (s.Length > maxLen)
            {
                s = s.Substring(0, maxLen).TrimEnd();
            }

            return s;
        }

        private static decimal? ParseAmount(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            var s = value.Trim();
            s = s.Replace(",", string.Empty);
            s = s.Replace("¥", string.Empty).Replace("￥", string.Empty).Replace("$", string.Empty);

            var neg = false;
            if (s.Length >= 2 && s[0] == '(' && s[s.Length - 1] == ')')
            {
                neg = true;
                s = s.Substring(1, s.Length - 2);
            }

            if (!decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                if (!decimal.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                {
                    return null;
                }
            }

            return neg ? -d : d;
        }

        private static Dictionary<string, decimal> BuildAmountByOrderId(List<Dictionary<string, string>> rows, string orderIdColumn, string amountColumn)
        {
            var dict = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                var oid = Get(row, orderIdColumn);
                if (string.IsNullOrWhiteSpace(oid)) continue;

                var amount = ParseAmount(Get(row, amountColumn));
                if (!amount.HasValue) continue;

                if (!dict.TryGetValue(oid, out var prev) || amount.Value > prev)
                {
                    dict[oid] = amount.Value;
                }
            }

            return dict;
        }

        private static string Get(Dictionary<string, string> row, string col)
        {
            if (row == null) return string.Empty;
            if (string.IsNullOrEmpty(col)) return string.Empty;

            return row.TryGetValue(col, out var v) ? (v ?? string.Empty).Trim() : string.Empty;
        }

        private sealed class CsvTable
        {
            public List<string> Headers { get; } = new List<string>();
            public List<Dictionary<string, string>> Rows { get; } = new List<Dictionary<string, string>>();
        }

        private static CsvTable ReadCsv(string path)
        {
            // try utf-8-sig then utf-8 then cp932
            foreach (var encoding in new[] { new UTF8Encoding(encoderShouldEmitUTF8Identifier: true), new UTF8Encoding(false), Encoding.GetEncoding(932) })
            {
                try
                {
                    using (var reader = new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true))
                    {
                        return ReadCsvWithReader(reader);
                    }
                }
                catch
                {
                    // try next
                }
            }

            // last resort
            using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                return ReadCsvWithReader(reader);
            }
        }

        private static CsvTable ReadCsvWithReader(TextReader reader)
        {
            // Basic CSV parser (handles quotes)
            var table = new CsvTable();

            var header = ReadCsvRecord(reader);
            if (header == null) return table;

            table.Headers.AddRange(header);

            while (true)
            {
                var rec = ReadCsvRecord(reader);
                if (rec == null) break;

                var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < table.Headers.Count; i++)
                {
                    var col = table.Headers[i];
                    var val = i < rec.Count ? rec[i] : string.Empty;
                    row[col] = val;
                }

                table.Rows.Add(row);
            }

            return table;
        }

        private static List<string> ReadCsvRecord(TextReader reader)
        {
            var line = reader.ReadLine();
            if (line == null) return null;

            var result = new List<string>();
            var sb = new StringBuilder();
            var inQuotes = false;

            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        // escaped quote
                        if (i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    if (c == ',')
                    {
                        result.Add(sb.ToString());
                        sb.Length = 0;
                    }
                    else if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
            }

            result.Add(sb.ToString());
            return result;
        }

        private static void WriteCsv(string path, List<Dictionary<string, string>> rows)
        {
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // UTF-8 BOM
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                writer.WriteLine(string.Join(",", OutputColumns.Select(EscapeCsv)));
                foreach (var row in rows)
                {
                    var values = OutputColumns.Select(c => EscapeCsv(row.TryGetValue(c, out var v) ? v : string.Empty));
                    writer.WriteLine(string.Join(",", values));
                }
            }
        }

        private static string EscapeCsv(string s)
        {
            s = s ?? string.Empty;
            if (s.Contains("\"") || s.Contains(",") || s.Contains("\n") || s.Contains("\r"))
            {
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            }

            return s;
        }
    }
}
