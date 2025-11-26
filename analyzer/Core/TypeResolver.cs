using System;
using System.Collections.Generic;

namespace analyzer.Core
{
    internal sealed class TypeResolver
    {
        private readonly IReadOnlyList<TypeKeyword> _keywords;

        public TypeResolver(IReadOnlyList<TypeKeyword> keywords)
        {
            _keywords = keywords ?? Array.Empty<TypeKeyword>();
        }

        public string Resolve(string target)
        {
            if (string.IsNullOrEmpty(target) || _keywords.Count == 0)
            {
                return null;
            }

            foreach (var entry in _keywords)
            {
                if (target.IndexOf(entry.Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return entry.Type;
                }
            }

            return null;
        }
    }
}

