using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Windows.Forms;

namespace HeaderMarkup
{
    using Table = List<RectArea>;
    using Sheet = Dictionary<RectArea, List<RectArea>>;
    using Book = Dictionary<string, Dictionary<RectArea, List<RectArea>>>;

    class RectArea
    {
        #region 构造器

        private struct Edge
        {
            public string edge;
            public Edge(string s) => edge = s;
            public Edge(Edge e) => this.edge = e.edge;
            public static int Compare(Edge lhs, Edge rhs)
            {
                char[] lhsc = lhs.edge.ToCharArray();
                char[] rhsc = rhs.edge.ToCharArray();
                if (lhsc.Length > rhsc.Length) return 1;
                else if (lhsc.Length < rhsc.Length) return -1;
                else
                    for (int i = 0; i < lhsc.Length; i++)
                        if (lhsc[i] > rhsc[i]) return 1;
                        else if (lhsc[i] < rhsc[i]) return -1;
                return 0;
            }
            public static bool operator >(Edge lhs, Edge rhs) => Compare(lhs, rhs) > 0;
            public static bool operator <(Edge lhs, Edge rhs) => Compare(lhs, rhs) < 0;
            public static bool operator >=(Edge lhs, Edge rhs) => Compare(lhs, rhs) >= 0;
            public static bool operator <=(Edge lhs, Edge rhs) => Compare(lhs, rhs) <= 0;
        }

        public string address { get; }
        private Edge left, top, right, bottom;
        private RectArea(string address)
        {
            this.address = address;
            string[] temp = address.Split(':', '$');
            int i = 0;
            foreach (var s in temp)
            {
                if (s.Length == 0) continue;
                if (i == 0) left = new Edge(s);
                else if (i == 1) top = new Edge(s);
                else if (i == 2) right = new Edge(s);
                else bottom = new Edge(s);
                i++;
            }
            if (i == 2)
            {
                right = new Edge(left);
                bottom = new Edge(top);
            }
        }

        public static RectArea GetArea(Excel.Range range)
        {
            if (!LegalRange(range)) return null;
            return new RectArea(range.Address);
        }

        private static bool LegalRange(Excel.Range range)
        {
            if (range == null) return false;
            if (range.Areas.Count != 1 || range.Height + range.Width > Properties.Settings.Default.MaxMarkupEdgeSize)
                return false;
            return true;
        }

        #endregion

        public bool IsOverlap(RectArea area) => !((this.left > area.right) || (this.right < area.left) || (this.top > area.bottom) || (this.bottom < area.top));

        public bool IsInside(RectArea area) => (this.left >= area.left) && (this.top >= area.top) && (this.right <= area.right) && (this.bottom <= area.bottom);

        public override string ToString() => ("[" + this.left.edge + this.top.edge + ":" + this.right.edge + this.bottom.edge + "]");
    }


    class Markups
    {
        public static Markups markups = new Markups();
        private Markups() => books = new Dictionary<string, Book>();

        private Dictionary<string, Book> books;

        private Sheet GetTables(Excel.Workbook workbook)
        {
            try
            {
                if(!(workbook.ActiveSheet is Excel.Worksheet)) return null;
                if (!books.ContainsKey(workbook.Name)) books.Add(workbook.Name, new Book());
                var sheets = books[workbook.Name];
                if (!sheets.ContainsKey(workbook.ActiveSheet.Name)) sheets.Add(workbook.ActiveSheet.Name, new Sheet());
                return sheets[workbook.ActiveSheet.Name];
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return null;
        }

        public void MarkTable(Excel.Workbook workbook, Excel.Range range)
        {
            RectArea area = RectArea.GetArea(range);
            Sheet tables = GetTables(workbook);
            if (area == null || tables == null) return;
            foreach (var item in tables)
                if (area.IsOverlap(item.Key)) return;
            tables.Add(area, new Table());
            DrawTable(range);
        }

        public void MarkHeader(Excel.Workbook workbook, Excel.Range range)
        {
            RectArea area = RectArea.GetArea(range);
            Sheet tables = GetTables(workbook);
            if (area == null || tables == null) return;
            Table table = null;
            foreach (var item in tables)
                if (area.IsInside(item.Key))
                {
                    table = item.Value;
                    break;
                }
            if (table == null) return;
            table.Add(area);
            DrawHeader(range);
        }

        public void EraseShapes(Excel.Worksheet worksheet)
        {
            try
            {
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    if (shape.Name.Contains(Properties.Settings.Default.MarkupTable) || shape.Name.Contains(Properties.Settings.Default.MarkupHeader))
                        shape.Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void RedrawShapes(Excel.Workbook workbook)
        {
            try
            {
                Sheet tables = GetTables(workbook);
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                EraseShapes(worksheet);
                if (tables == null) return;
                foreach (var table in tables)
                {
                    DrawTable(worksheet.Range[table.Key.address]);
                    foreach (var header in table.Value)
                        DrawHeader(worksheet.Range[header.address]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void Reset(Excel.Workbook workbook)
        {
            try
            {
                Sheet tables = GetTables(workbook);
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                EraseShapes(worksheet);
                while(tables.Count != 0)
                    tables.Remove(tables.Keys.FirstOrDefault());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GenString(Excel.Workbook workbook)
        {
            if (!books.ContainsKey(workbook.Name)) return "";
            Book book = books[workbook.Name];
            string temp = "";
            foreach (var sheet in book)
            {
                temp += "[" + sheet.Key;
                foreach (var table in sheet.Value)
                {
                    temp += "[" + table.Key.ToString();
                    foreach (var header in table.Value)
                        temp += header.ToString();
                    temp += "]";
                }
                temp += "]";
            }
            return temp;
        }

        public string SaveMarkup(Excel.Workbook workbook, bool saveShapes, bool saveProperty)
        {
            if (!saveShapes)
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    EraseShapes(worksheet);
            string markupProperty = GenString(workbook);
            if (saveProperty && markupProperty.Length < 1024)
            {
                Office.DocumentProperties properties = workbook.CustomDocumentProperties;
                foreach (Office.DocumentProperty property in properties)
                    if (property.Name == "HMarker")
                        property.Delete();
                properties.Add("HMarker", false, Office.MsoDocProperties.msoPropertyTypeString, markupProperty, null);
            }
            return markupProperty;
        }

        public void Remove(Excel.Workbook workbook) => books.Remove(workbook.Name);

        #region 画图

        // 图形标注
        private static float[,] TableLine(double x1, double y1, double x2, double y2)
        {
            int steps = Math.Abs((int)((x2 - x1 + y2 - y1) / 6)) / 2 * 2 + 1;
            double start = 0, step = 1.0 / steps / 3;
            double dhx = (y2 - y1) * step * 3 / 4, dhy = -(x2 - x1) * step * 3 / 4;
            float[,] polyArray = new float[steps * 2 + 2, 2];
            polyArray[0, 0] = Convert.ToSingle(x1);
            polyArray[0, 1] = Convert.ToSingle(y1);
            for (int i = 0; i < steps; i++)
            {
                for (int j = 1; j < 3; j++)
                {
                    start += step;
                    polyArray[2 * i + j, 0] = Convert.ToSingle(Math.Max(x1 * (1 - start) + x2 * start + dhx, 0));
                    polyArray[2 * i + j, 1] = Convert.ToSingle(Math.Max(y1 * (1 - start) + y2 * start + dhy, 0));
                }
                start += step;
                dhx = -dhx;
                dhy = -dhy;
            }
            polyArray[steps * 2 + 1, 0] = Convert.ToSingle(x2);
            polyArray[steps * 2 + 1, 1] = Convert.ToSingle(y2);
            return polyArray;
        }

        private static void DrawTable(Excel.Range range)
        {
            Excel.Worksheet worksheet = range.Worksheet;
            double x = range.Left, y = range.Top, w = range.Width, h = range.Height;
            object[] shapes = new object[4];
            double[,] points = new double[4, 2] { { x, y }, { x + w, y }, { x + w, y + h }, { x, y + h } };
            for (int i = 0; i < 4; i++)
            {
                Excel.Shape line = worksheet.Shapes.AddPolyline(TableLine(points[i, 0], points[i, 1], points[(i + 1) % 4, 0], points[(i + 1) % 4, 1]));
                line.Name = Properties.Settings.Default.MarkupTableLine + Properties.Settings.Default.MarkupShapesCount.ToString() + "-" + i.ToString();
                shapes[i] = line.Name;
            }
            Excel.ShapeRange shapeRange = worksheet.Shapes.Range[shapes];
            shapeRange.Group();
            shapeRange.Name = Properties.Settings.Default.MarkupTable + Properties.Settings.Default.MarkupShapesCount.ToString();
            Properties.Settings.Default.MarkupShapesCount++;
        }

        private struct Point
        {
            public float x, y;
            public void SetPoint(double x, double y)
            {
                this.x = Convert.ToSingle(x);
                this.y = Convert.ToSingle(y);
            }
        }

        private static void DrawHeader(Excel.Range range)
        {
            Excel.Worksheet worksheet = range.Worksheet;
            double x = range.Left, y = range.Top, w = range.Width, h = range.Height;
            int start = (int)((x + y) / 6) + 1, end = (int)((x + y + h + w) / 6);
            object[] lines = new object[end - start + 1];
            Point point1 = new Point(), point2 = new Point();
            for (int i = start; i <= end; i++)
            {
                if (6 * i <= x + y + w)
                    point1.SetPoint(6 * i - y, y);
                else
                    point1.SetPoint(x + w, 6 * i - x - w);
                if (6 * i <= x + y + h)
                    point2.SetPoint(x, 6 * i - x);
                else
                    point2.SetPoint(6 * i - y - h, y + h);
                var line = worksheet.Shapes.AddLine(point1.x, point1.y, point2.x, point2.y);
                line.Name = Properties.Settings.Default.MarkupHeaderLine + Properties.Settings.Default.MarkupShapesCount.ToString() + "_" + i.ToString();
                lines[i - start] = line.Name;
            }
            Excel.ShapeRange shapeRange = worksheet.Shapes.Range[lines];
            shapeRange.Group();
            shapeRange.Name = Properties.Settings.Default.MarkupHeader + Properties.Settings.Default.MarkupShapesCount.ToString();
            Properties.Settings.Default.MarkupShapesCount++;
        }

        #endregion

    }
}
