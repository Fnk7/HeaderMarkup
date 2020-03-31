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
    using Sheet = List<RectArea>;
    using Book = Dictionary<string, List<RectArea>>;

    #region RectArea 区域的实现

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

        public string Address { get; }
        private Edge left, top, right, bottom;

        protected RectArea(string address)
        {
            MessageBox.Show(address);
            this.Address = address.Trim();
            string[] temp = Address.Split(':', '$');
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

        protected static bool LegalRange(Excel.Range range)
        {
            if (range == null) return false;
            if (range.Areas.Count != 1 || range.Height + range.Width > Share.settings.MaxEdgeSize)
                return false;
            return true;
        }

        #endregion

        public static RectArea GetRectArea(Excel.Range range)
        {
            if (!RectArea.LegalRange(range)) return null;
            return new RectArea(range.Address);
        }

        public bool IsOverlap(RectArea area) => !((this.left > area.right) || (this.right < area.left) || (this.top > area.bottom) || (this.bottom < area.top));

        public bool IsInside(RectArea area) => (this.left >= area.left) && (this.top >= area.top) && (this.right <= area.right) && (this.bottom <= area.bottom);

    }

    class Table : RectArea
    {
        public List<MarkArea> MarkAreas { get; }
        private Table(string address) : base(address) => MarkAreas = new List<MarkArea>();
        public static Table GetTable(Excel.Range range)
        {
            if (!RectArea.LegalRange(range)) return null;
            return new Table(range.Address);
        }
        public void AddMarkArea(MarkArea markArea) => MarkAreas.Add(markArea);
        public override string ToString()
        {
            string temp = "[Tb" + Address;
            foreach (var markArea in MarkAreas)
                temp += markArea.ToString();
            temp += "]";
            return temp;
        }
    }

    class MarkArea : RectArea
    {
        // 0 代表确认是表头，1，2代表偏向数据的程度，-1，-2偏向标题的程度。
        public int Type { get; }
        public static readonly int LikeTitle2 = -2, LikeTitle1 = -1, Header = 0, LikeData1 = 1, LikeData2 = 2;
        private MarkArea(string address, int type) : base(address) => this.Type = type;
        public static MarkArea GetMarkArea(Excel.Range range, int type)
        {
            if (!RectArea.LegalRange(range)) return null;
            return new MarkArea(range.Address, type);
        }
        public override string ToString() => "[Mk" + Type.ToString() + Address + "]";
    }

    #endregion

    class Markups
    {
        public Markups() => books = new Dictionary<string, Book>();

        private Dictionary<string, Book> books;


        #region 标注

        private Sheet GetSheet(Excel.Workbook workbook)
        {
            try
            {
                if (!(workbook.ActiveSheet is Excel.Worksheet)) return null;
                if (!books.ContainsKey(workbook.Name)) books.Add(workbook.Name, new Book());
                var sheets = books[workbook.Name];
                if (!sheets.ContainsKey(workbook.ActiveSheet.Name)) sheets.Add(workbook.ActiveSheet.Name, new Sheet());
                return sheets[workbook.ActiveSheet.Name];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return null;
        }

        public void AddTable(Excel.Workbook workbook, Excel.Range range)
        {
            Table table = Table.GetTable(range);
            Sheet sheet = GetSheet(workbook);
            if (table == null || sheet == null) return;
            foreach (var rectArea in sheet)
                if (rectArea is Table && table.IsOverlap(rectArea)) return;
            sheet.Add(table);
            DrawTable(range);
        }

        public void AddMarkArea(Excel.Workbook workbook, Excel.Range range, int type)
        {
            MarkArea markArea = MarkArea.GetMarkArea(range, type);
            Sheet sheet = GetSheet(workbook);
            if (markArea == null || sheet == null) return;
            Table table = null;
            foreach (var rectArea in sheet)
            {
                if (rectArea is Table)
                {
                    foreach (var temp in ((Table)rectArea).MarkAreas)
                        if (markArea.IsOverlap(temp))
                            return;
                    if (markArea.IsInside(rectArea))
                        table = (Table)rectArea;
                }
                else if (markArea.IsOverlap(rectArea))
                    return;
            }
            if (type == MarkArea.LikeTitle2)
                sheet.Add(markArea);
            else if (table == null) return;
            else
                table.AddMarkArea(markArea);
            DrawMarkArea(range, type);
        }

        #endregion

        #region 删除

        private void EraseShapes(Excel.Worksheet worksheet)
        {
            try
            {
                foreach (Excel.Shape shape in worksheet.Shapes)
                {
                    if (shape.Name.Contains(Share.settings.TableShapeName) || shape.Name.Contains(Share.settings.MarkAreaName))
                        shape.Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void RedrawShapes(Excel.Workbook workbook)
        {
            try
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                EraseShapes(worksheet);
                Sheet sheet = GetSheet(workbook);
                if (sheet == null) return;
                foreach (var rectArea in sheet)
                {
                    if (rectArea is Table)
                    {
                        Table table = (Table)rectArea;
                        DrawTable(worksheet.Range[table.Address]);
                        foreach (var markArea in table.MarkAreas)
                            DrawMarkArea(worksheet.Range[markArea.Address], markArea.Type);
                    }
                    else
                        DrawMarkArea(worksheet.Range[rectArea.Address], ((MarkArea)rectArea).Type);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void DeleteAll(Excel.Workbook workbook)
        {
            try
            {
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                EraseShapes(worksheet);
                Sheet sheet = GetSheet(workbook);
                while (sheet.Count != 0)
                    sheet.Remove(sheet.FirstOrDefault());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void DeleteArea(Excel.Workbook workbook, Excel.Range range)
        {
            RectArea rectArea = RectArea.GetRectArea(range);
            Sheet sheet = GetSheet(workbook);
            if (rectArea == null || sheet == null) return;
            RectArea rectTemp = null;
            foreach (var rect in sheet)
            {
                if (rect is Table)
                {
                    MarkArea markTemp = null;
                    foreach (var mark in ((Table)rect).MarkAreas)
                        if (rectArea.IsInside(mark))
                        {
                            markTemp = mark;
                            break;
                        }
                    if (markTemp != null)
                    {
                        ((Table)rect).MarkAreas.Remove(markTemp);
                        break;
                    }
                }
                else if (rectArea.IsInside(rect))
                {
                    rectTemp = rect;
                    break;
                }
            }
            if (rectTemp != null)
                sheet.Remove(rectTemp);
            RedrawShapes(workbook);
        }

        public void DeleteTable(Excel.Workbook workbook, Excel.Range range)
        {
            RectArea rectArea = RectArea.GetRectArea(range);
            Sheet sheet = GetSheet(workbook);
            if (rectArea == null || sheet == null) return;
            RectArea table = null;
            foreach (var rect in sheet)
                if (rect is Table && rectArea.IsInside(rect))
                {
                    table = rect;
                    break;
                }
            if (table != null)
                sheet.Remove(table);
            RedrawShapes(workbook);
        }

        #endregion

        #region 保存

        private string GenString(Excel.Workbook workbook)
        {
            if (!books.ContainsKey(workbook.Name)) return "";
            Book book = books[workbook.Name];
            string temp = "";
            foreach (var sheet in book)
            {
                temp += "[" + sheet.Key;
                foreach (var rectArea in sheet.Value)
                    temp += rectArea.ToString();
                temp += "]";
            }
            return temp;
        }

        public string MarkupInfos(Excel.Workbook workbook)
        {
            if (!Share.settings.SaveMarkupShapes)
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    EraseShapes(worksheet);
            string markupProperty = GenString(workbook);
            if (Share.settings.SaveInWorkbookProperty && markupProperty.Length < 1024)
            {
                Office.DocumentProperties properties = workbook.CustomDocumentProperties;
                foreach (Office.DocumentProperty property in properties)
                    if (property.Name == "HMarkup")
                        property.Delete();
                properties.Add("HMarkup", false, Office.MsoDocProperties.msoPropertyTypeString, markupProperty, null);
            }
            return markupProperty;
        }

        public void Remove(Excel.Workbook workbook) => books.Remove(workbook.Name);

        #endregion

        #region 画图
        // Line.ForeColor.RGB 和 color.TOArgb红蓝位置相反
        private static int RGBColor(System.Drawing.Color color) => color.B * 0x10000 + color.G * 0x100 + color.R;

        // Table的一条边界
        private static float[,] TableEdge(double x1, double y1, double x2, double y2)
        {
            float interval = Share.settings.TableInterval;
            int steps = Math.Abs((int)((x2 - x1 + y2 - y1) / interval)) / 2 * 2 + 1;
            float[,] polyArray = new float[steps * 2 + 2, 2];
            polyArray[0, 0] = Convert.ToSingle(x1);
            polyArray[0, 1] = Convert.ToSingle(y1);
            double start = 0, step = 1.0 / steps / 3;
            double dhx = (y2 - y1) * step * 3 / 4, dhy = -(x2 - x1) * step * 3 / 4;
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

        // 画出table
        private static void DrawTable(Excel.Range range)
        {
            Excel.Worksheet worksheet = range.Worksheet;
            double left = range.Left, top = range.Top, width = range.Width, height = range.Height;
            object[] shapes = new object[4];
            double[,] points = new double[4, 2] { { left, top }, { left + width, top }, { left + width, top + height }, { left, top + height } };
            for (int i = 0; i < 4; i++)
            {
                Excel.Shape edge = worksheet.Shapes.AddPolyline(TableEdge(points[i, 0], points[i, 1], points[(i + 1) % 4, 0], points[(i + 1) % 4, 1]));
                edge.Name = Share.settings.TableEdgeName + Share.settings.ShapeCount.ToString() + "-" + i.ToString();
                shapes[i] = edge.Name;
            }
            Excel.Shape shape = worksheet.Shapes.Range[shapes].Group();
            shape.Line.Weight = 2.0f;
            shape.Line.ForeColor.RGB = RGBColor(System.Drawing.Color.Blue);
            shape.Name = Share.settings.TableShapeName + Share.settings.ShapeCount.ToString();
            Share.settings.ShapeCount++;
        }

        // 画出MarkArea的阴影
        private struct Point
        {
            public float x, y;
            public void SetPoint(double x, double y)
            {
                this.x = Convert.ToSingle(x);
                this.y = Convert.ToSingle(y);
            }
        }

        private static void DrawMarkArea(Excel.Range range, int type)
        {
            Excel.Worksheet worksheet = range.Worksheet;
            double left = range.Left, top = range.Top, width = range.Width, height = range.Height;
            float interval = Share.settings.MarkInterval;
            int start = (int)((left + top) / interval) + 1, end = (int)((left + top + height + width) / interval);
            object[] lines = new object[end - start + 1];
            Point point1 = new Point(), point2 = new Point();
            for (int i = start; i <= end; i++)
            {
                if (interval * i <= left + top + width)
                    point1.SetPoint(interval * i - top, top);
                else
                    point1.SetPoint(left + width, interval * i - left - width);
                if (interval * i <= left + top + height)
                    point2.SetPoint(left, interval * i - left);
                else
                    point2.SetPoint(interval * i - top - height, top + height);
                var line = worksheet.Shapes.AddLine(point1.x, point1.y, point2.x, point2.y);
                line.Name = Share.settings.MarkLineName + Share.settings.ShapeCount.ToString() + "-" + i.ToString();
                lines[i - start] = line.Name;
            }
            Excel.Shape shape = worksheet.Shapes.Range[lines].Group();
            shape.Name = Share.settings.MarkAreaName + Share.settings.ShapeCount.ToString();
            shape.Line.Weight = 1.5f;
            shape.Line.ForeColor.RGB = RGBColor(System.Drawing.Color.FromArgb(Math.Min(255, 255 + type * 96), Math.Max(0, type * 96), 31));
            Share.settings.ShapeCount++;
        }

        #endregion

    }
}
