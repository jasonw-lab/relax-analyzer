namespace analyzer.Core
{
    internal sealed class TypeKeyword
    {
        public TypeKeyword(string keyword, string type)
        {
            Keyword = keyword;
            Type = type;
        }

        public string Keyword { get; }

        public string Type { get; }
    }
}

