using System.Collections.Generic;

namespace analyzer.Core
{
    internal sealed class SheetRow
    {
        public SheetRow(string[] values, bool isFileHeader)
        {
            Values = values;
            IsFileHeader = isFileHeader;
        }

        public string[] Values { get; }

        public bool IsFileHeader { get; }
    }

    internal sealed class FileImportChunk
    {
        public FileImportChunk(string fileName)
        {
            FileName = fileName;
            Rows = new List<SheetRow>();
        }

        public string FileName { get; }

        public List<SheetRow> Rows { get; }
    }

    internal sealed class SheetImportBatch
    {
        public SheetImportBatch(string sheetName)
        {
            SheetName = sheetName;
            Chunks = new List<FileImportChunk>();
        }

        public string SheetName { get; }

        public List<FileImportChunk> Chunks { get; }
    }
}

