using System;
using System.Linq;
using System.Collections.Generic;

namespace HeaderMarkup.Markup
{
    class MarkSheet
    {
        public List<MarkRange> ranges = new List<MarkRange>();

        public void AddTable(string address)
        {
            MarkTable table = new MarkTable(address);

            foreach (var range in ranges)
                if (range is MarkTable && range.IsOverlap(table))
                    throw new Exception($"New Table[{address}] Overlaps Table[{range.Address}].");
            ranges.Add(table);
        }

        public void AddHeader(string address, int type)
        {
            MarkHeader header = new MarkHeader(address, type);
            MarkTable table = null;
            foreach (var range in ranges)
            {
                if (range is MarkHeader)
                    if (range.IsOverlap(header))
                        throw new Exception($"New Header[{address}] Overlaps Header[{range.Address}].");
                    else continue;
                if (((MarkTable)range).headers.Any(temp => temp.IsOverlap(header)))
                    throw new Exception($"New Header[{address}] Overlaps Header in Table[{range.Address}].");
                if (header.IsInside(range))
                    table = (MarkTable)range;
            }
            if (type == -2)
                ranges.Add(header);
            else if(table != null) 
                table.headers.Add(header);
            else
                throw new Exception($"New Header[{address}] is not Inside any Table.");
            return;
        }

        public string DeletAll()
        {
            var delete = string.Concat(ranges);
            ranges.Clear();
            return delete;
        }

        public string DeletTable(string address)
        {
            var range = new MarkRange(address);
            foreach (var temp in ranges.OfType<MarkTable>())
            {
                if (!range.IsInside(temp))
                    continue;
                ranges.Remove(temp);
                return temp.ToString();
            }
            return string.Empty;
        }

        public string DeletHeader(string address)
        {
            var temp = new MarkRange(address);
            foreach (var range in ranges.Where(r => temp.IsInside(r)))
            {
                if (range is MarkHeader)
                {
                    ranges.Remove(range);
                    return range.ToString();
                }
                foreach (var header in ((MarkTable)range).headers.Where(h => temp.IsInside(h)))
                {
                    ((MarkTable)range).headers.Remove(header);
                    return header.ToString();
                }
            }
            return string.Empty;
        }

        public override string ToString()
            => string.Concat(ranges);
    }
}
