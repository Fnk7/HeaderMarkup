using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace HeaderMarkup.DrawShape
{
    class EraseShape
    {
        public static void EraseAll()
        {
            Excel.Workbook workbook = Utils.GetActiveWorkbook();
            Excel.Worksheet worksheet = Utils.GetActiveWorksheet(workbook);
            foreach (Excel.Shape shape in worksheet.Shapes)
                if (shape.Name.Contains(Share.settings.TableShapeName) || shape.Name.Contains(Share.settings.HeaderShapeName))
                    shape.Delete();
        }

        public static void EraseByName(Excel.Worksheet worksheet, string name)
        {
            foreach (Excel.Shape shape in worksheet.Shapes)
                if (shape.Name == Share.settings.TableShapeName + name || shape.Name == Share.settings.HeaderShapeName + name)
                {
                    shape.Delete();
                    break;
                }
        }

        public static void EraseByName(string name)
        {
            Excel.Workbook workbook = Utils.GetActiveWorkbook();
            Excel.Worksheet worksheet = Utils.GetActiveWorksheet(workbook);
            EraseByName(worksheet, name);
        }

        public static void EraseByName(List<string> names)
        {
            Excel.Workbook workbook = Utils.GetActiveWorkbook();
            Excel.Worksheet worksheet = Utils.GetActiveWorksheet(workbook);
            foreach (var name in names)
                EraseByName(worksheet, name);
        }
    }
}
