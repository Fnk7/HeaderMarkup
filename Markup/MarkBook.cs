using System.Linq;
using System.Collections.Generic;

namespace HeaderMarkup.Markup
{
    class MarkBook
    {
        public Dictionary<string, MarkSheet> markSheets = new Dictionary<string, MarkSheet>();
        public override string ToString()
            => string.Concat(markSheets.Select(keyValue => $"{keyValue.Key}{keyValue.Value.ToString()}\n"));
    }
}
