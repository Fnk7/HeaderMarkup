using System;
using System.Linq;
using System.Collections.Generic;


namespace HeaderMarkup.Markup
{
    class Sheet
    {
        public List<Range> ranges = new List<Range>();

        public string AddTable(string address)
        {
            Table table = new Table(address);
            foreach (var range in ranges)
                if (range.IsOverlap(table))
                    throw new Exception($"New Table[{address}] Overlaps Title/Table[{range.Address}].");
            ranges.Add(table);
            return table.Name;
        }

        public string AddMark(string address, int type)
        {
            Mark mark = new Mark(address, type);
            if (type == -2)
            {
                foreach (var range in ranges)
                    if (range.IsOverlap(mark))
                        throw new Exception($"New Title[{address}] Overlaps Title/Table[{range.Address}].");
                ranges.Add(mark);
                return mark.Name;
            }
            Table table = ranges.OfType<Table>().FirstOrDefault(t =>
            {
                if (!mark.IsInside(t))
                    return false;
                foreach (var m in t.marks)
                    if (m.IsOverlap(mark))
                        throw new Exception($"New Mark[{address}] Overlaps Mark[{m.Address}].");
                return true;
            });
            if (table == null)
                throw new Exception($"New Mark[{address}] is not Inside any Table.");
            table.marks.Add(mark);
            return mark.Name;
        }

        public void DeletAll() => ranges.Clear();

        public List<string> DeletTable(string address)
        {
            var range = new Range(address);
            var toDelete = new List<string>();
            var table = ranges.OfType<Table>().FirstOrDefault(t => range.IsInside(t));
            if (table != null)
            {
                ranges.Remove(table);
                toDelete.Add(table.Name);
                toDelete = toDelete.Concat(table.marks.Select(mark => mark.Name)).ToList();
            }
            return toDelete;
        }

        public string DeletMark(string address)
        {
            var range = new Range(address);
            var temp = ranges.FirstOrDefault(r => range.IsInside(r));
            if (temp is Mark mark)
            {
                ranges.Remove(mark);
                return mark.Name;
            }
            if (temp is Table table)
            {
                mark = table.marks.FirstOrDefault(m => range.IsInside(m));
                if (mark != null)
                {
                    table.marks.Remove(mark);
                    return mark.Name;
                }
            }
            return string.Empty;
        }

        public override string ToString()
            => string.Concat(ranges);
    }
}
