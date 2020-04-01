using System;

using Excel = Microsoft.Office.Interop.Excel;

namespace HeaderMarkup.DrawShape
{
    public struct Point
    {
        public float x, y;
        public void Set(float x, float y)
        {
            this.x = x;
            this.y = y;
        }
    }

    class DrawHeader
    {
        public static void Draw(Excel.Range range, int type, string name)
        {
            var worksheet = range.Worksheet;
            var interval = Share.settings.HeaderInterval;
            float left = (float)range.Left, top = (float)range.Top;
            float width = (float)range.Width, height = (float)range.Height;
            int start = (int)((left + top) / interval) + 1;
            int end = (int)((left + top + height + width) / interval);
            string[] lines = new string[end - start + 1];
            Point p1 = new Point(), p2 = new Point();
            for (int current = start; current <= end; current++)
            {
                if (interval * current <= left + top + width)
                    p1.Set(interval * current - top, top);
                else
                    p1.Set(left + width, interval * current - left - width);
                if (interval * current <= left + top + height)
                    p2.Set(left, interval * current - left);
                else
                    p2.Set(interval * current - top - height, top + height);
                var line = worksheet.Shapes.AddLine(p1.x, p1.y, p2.x, p2.y);
                line.Name = Share.settings.HeaderLineName + name + current;
                lines[current - start] = line.Name;
            }
            Excel.Shape headerShape = worksheet.Shapes.Range[lines].Group();
            headerShape.Name = Share.settings.HeaderShapeName + name;
            headerShape.Line.Weight = Share.settings.HeaderLineWeight;
            headerShape.Line.ForeColor.RGB = Utils.RGBColor(System.Drawing.Color.FromArgb(Math.Min(255, 255 + type * 96), Math.Max(0, type * 96), 31));
        }
    }
}
