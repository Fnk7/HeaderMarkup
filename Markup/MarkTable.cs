using System;
using System.Collections.Generic;

namespace HeaderMarkup.Markup
{
    class MarkTable : MarkRange
    {
        public List<MarkHeader> headers;
        public MarkTable(string address) : base(address)
        {
            if (left == right || top == bottom)
                throw new Exception($"Invalid Table: {address}.");
            headers = new List<MarkHeader>();
        }

        public override string ToString()
            => $"[Tb,{headers.Count},{left},{top},{right},{bottom}]"
            + $"{string.Concat(headers)}";
    }
}
