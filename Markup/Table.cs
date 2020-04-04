using System;
using System.Collections.Generic;

namespace HeaderMarkup.Markup
{
    class Table : Range
    {
        public List<Mark> marks;
        public Table(string address) : base(address)
        {
            if (left == right || top == bottom)
                throw new Exception($"Invalid Table: {address}.");
            marks = new List<Mark>();
        }

        public override string Name
            => $"[Tb,{left},{top},{right},{bottom}]";

        public override string ToString()
            => $"[Tb,{marks.Count},{left},{top},{right},{bottom}]"
            + $"{string.Concat(marks)}";
    }
}
