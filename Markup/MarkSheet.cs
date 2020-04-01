using System;
using System.Linq;
using System.Collections.Generic;


namespace HeaderMarkup.Markup
{
    class MarkSheet
    {
        public List<MarkRange> ranges = new List<MarkRange>();

        public string AddTable(string address)
        {
            MarkTable table = new MarkTable(address);

            foreach (var range in ranges)
                if (range is MarkTable && range.IsOverlap(table))
                    throw new Exception($"New Table[{address}] Overlaps Table[{range.Address}].");
            ranges.Add(table);
            return table.Name;
        }

        public string AddHeader(string address, int type)
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
            return header.Name;
        }

        public void DeletAll() => ranges.Clear();

        public List<string> DeletTable(string address)
        {
            var range = new MarkRange(address);
            var toDelete = new List<string>();
            var table = ranges.OfType<MarkTable>().FirstOrDefault(t => range.IsInside(t));
            if (table != null)
            {
                ranges.Remove(table);
                toDelete.Add(table.Name);
                toDelete = toDelete.Concat(table.headers.Select(header => header.Name)).ToList();
            }
            return toDelete;
        }

        public string DeletHeader(string address)
        {
            var range = new MarkRange(address);
            foreach (var temp in ranges.Where(r => range.IsInside(r)))
            {
                if (temp is MarkHeader)
                {
                    ranges.Remove(temp);
                    return temp.Name;
                }
                foreach (var header in ((MarkTable)temp).headers.Where(h => range.IsInside(h)))
                {
                    ((MarkTable)temp).headers.Remove(header);
                    return header.Name;
                }
            }
            return string.Empty;
        }

        public override string ToString()
            => string.Concat(ranges);
    }
}
