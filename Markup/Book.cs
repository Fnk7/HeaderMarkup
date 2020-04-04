using System.Linq;
using System.Collections.Generic;

namespace HeaderMarkup.Markup
{
    class Book
    {
        public Dictionary<string, Sheet> sheets = new Dictionary<string, Sheet>();
        public override string ToString()
            => string.Concat(sheets.Select(keyValue => $"{keyValue.Key}{keyValue.Value.ToString()}\n"));
    }
}
