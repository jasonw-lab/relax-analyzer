using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace analyzer.Core
{
    internal sealed class CsvImportService
    {
        private readonly TypeResolver _typeResolver;
        private readonly IList<string> _warnings;

        public CsvImportService(TypeResolver typeResolver, IList<string> warnings)
        {
            _typeResolver = typeResolver ?? new TypeResolver(Array.Empty<TypeKeyword>());
            _warnings = warnings ?? new List<string>();
        }

        public IReadOnlyList<SheetImportBatch> LoadFiles(IEnumerable<string> filePaths)
        {
            var batches = new Dictionary<string, SheetImportBatch>(StringComparer.OrdinalIgnoreCase);
            if (filePaths == null)
            {
                return Array.Empty<SheetImportBatch>();
            }

            foreach (var path in filePaths)
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    _warnings.Add($"ファイルが見つかりません: {path}");
                    continue;
                }

                if (!MonthExtractor.TryGetMonthSheet(path, out var sheetName, out var message))
                {
                    _warnings.Add(message);
                    continue;
                }

                var chunk = BuildChunkForFile(path);
                if (chunk == null || chunk.Rows.Count == 0)
                {
                    continue;
                }

                if (!batches.TryGetValue(sheetName, out var batch))
                {
                    batch = new SheetImportBatch(sheetName);
                    batches.Add(sheetName, batch);
                }

                batch.Chunks.Add(chunk);
            }

            return batches.Values.ToList();
        }

        private FileImportChunk BuildChunkForFile(string filePath)
        {
            var chunk = new FileImportChunk(Path.GetFileName(filePath));
            chunk.Rows.Add(new SheetRow(CreateFileHeaderRow(chunk.FileName), true));

            foreach (var record in ReadCsvRecords(filePath))
            {
                chunk.Rows.Add(new SheetRow(record, false));
            }

            if (chunk.Rows.Count == 1)
            {
                _warnings.Add($"データ行が存在しません: {filePath}");
            }

            return chunk;
        }

        private IEnumerable<string[]> ReadCsvRecords(string filePath)
        {
            var encodings = new[]
            {
                new EncodingInfo("UTF-8", new UTF8Encoding(false, true)),
                new EncodingInfo("Shift_JIS", Encoding.GetEncoding(932))
            };

            Exception lastError = null;
            foreach (var encoding in encodings)
            {
                List<string[]> records;
                try
                {
                    records = ReadWithEncoding(filePath, encoding.Value).ToList();
                }
                catch (Exception ex)
                {
                    lastError = ex;
                    continue;
                }

                foreach (var record in records)
                {
                    yield return record;
                }

                if (lastError != null)
                {
                    _warnings.Add($"文字コードを {encoding.Name} にフォールバックしました: {Path.GetFileName(filePath)}");
                }

                yield break;
            }

            throw lastError ?? new InvalidOperationException("CSV の読み込みに失敗しました。");
        }

        private IEnumerable<string[]> ReadWithEncoding(string filePath, Encoding encoding)
        {
            var cfg = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
                BadDataFound = null,
                DetectDelimiter = false,
                Delimiter = ","
            };

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
            using (var reader = new StreamReader(stream, encoding, true))
            using (var csv = new CsvReader(reader, cfg))
            {
                if (csv.Read())
                {
                    csv.ReadHeader();
                }

                while (csv.Read())
                {
                    yield return NormalizeRecord(csv);
                }
            }
        }

        private string[] NormalizeRecord(CsvReader csv)
        {
            var record = new string[12];
            for (var i = 0; i < record.Length; i++)
            {
                record[i] = string.Empty;
            }

            for (var i = 0; i < record.Length; i++)
            {
                string value;
                if (!csv.TryGetField(i, out value))
                {
                    break;
                }

                record[i] = value ?? string.Empty;
            }

            var storeName = record.Length > 1 ? record[1] : string.Empty;
            var resolvedType = _typeResolver.Resolve(storeName);
            if (!string.IsNullOrEmpty(resolvedType) && record.Length > 10)
            {
                record[10] = resolvedType;
            }

            return record;
        }

        private static string[] CreateFileHeaderRow(string fileName)
        {
            var row = new string[12];
            row[0] = fileName;
            return row;
        }

        private sealed class EncodingInfo
        {
            public EncodingInfo(string name, Encoding value)
            {
                Name = name;
                Value = value;
            }

            public string Name { get; }

            public Encoding Value { get; }
        }
    }
}

