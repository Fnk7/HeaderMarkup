using System;

using Excel = Microsoft.Office.Interop.Excel;

namespace HeaderMarkup.DrawShape
{
    class DrawTable
    {
        public static float[,] GetEdge(float x1, float y1, float x2, float y2)
        {
            float interval = Share.settings.TableInterval;
            int steps = Math.Abs((int)((x2 - x1 + y2 - y1) / interval)) / 2 * 2 + 1;
            float[,] edge = new float[steps * 2 + 2, 2];
            edge[0, 0] = x1;
            edge[0, 1] = y1;
            float start = 0, step = 1.0f / steps / 3;
            float dhx = (y2 - y1) * step * 3 / 4, dhy = -(x2 - x1) * step * 3 / 4;
            for (int i = 0; i < steps; i++)
            {
                for (int j = 1; j < 3; j++)
                {
                    start += step;
                    edge[2 * i + j, 0] = Math.Max(x1 * (1 - start) + x2 * start + dhx, 0);
                    edge[2 * i + j, 1] = Math.Max(y1 * (1 - start) + y2 * start + dhy, 0);
                }
                start += step;
                dhx = -dhx;
                dhy = -dhy;
            }
            edge[steps * 2 + 1, 0] = x2;
            edge[steps * 2 + 1, 1] = y2;
            return edge;
        }

        public static void Draw(Excel.Range range, string name)
        {
            Excel.Worksheet worksheet = range.Worksheet;
            float left = (float)range.Left, top = (float)range.Top;
            float width = (float)range.Width, height = (float)range.Height;
            string[] edges = new string[4];
            Point[] points = { 
                new Point { x = left, y = top},
                new Point { x = left + width, y = top},
                new Point { x = left + width, y = top + height},
                new Point { x = left, y = top + height}
            };
            for (int i = 0; i < 4; i++)
            {
                Excel.Shape edge = worksheet.Shapes.AddPolyline(GetEdge(points[i].x, points[i].y, points[(i + 1) % 4].x, points[(i + 1) % 4].y));
                edge.Name = Share.settings.TableEdgeName + name + i;
                edges[i] = edge.Name;
            }
            Excel.Shape tableShape = worksheet.Shapes.Range[edges].Group();
            tableShape.Name = Share.settings.TableShapeName + name;
            tableShape.Line.Weight = Share.settings.TableLineWeight;
            tableShape.Line.ForeColor.RGB = Utils.RGBColor(System.Drawing.Color.Blue);
        }
    }
}
