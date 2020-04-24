using Excel = Microsoft.Office.Interop.Excel;

namespace HeaderMarkup.DrawShape
{
    class DrawPredict
    {
        public static void Draw(Excel.Range range, string name)
        {
            var worksheet = range.Worksheet;
            var interval = Share.settings.MarkInterval;
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
                line.Name = Share.settings.PredictLineName + name + current;
                lines[current - start] = line.Name;
            }
            Excel.Shape markShape = worksheet.Shapes.Range[lines].Group();
            markShape.Name = Share.settings.PredictShapeName + name;
            markShape.Line.Weight = Share.settings.MarkLineWeight;
            markShape.Line.ForeColor.RGB = Utils.RGBColor(System.Drawing.Color.BlueViolet);
        }
    }
}
