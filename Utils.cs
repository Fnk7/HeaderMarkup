using System;
using System.Drawing;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

using HeaderMarkup.Setting;
using HeaderMarkup.Markup;

namespace HeaderMarkup
{
    static class Share
    {
        public static readonly string defualtDataset = "D:\\Temp\\Markup";
        public static readonly string defualtCSV = "D:\\Temp\\CSV";
        public static readonly string modelName = "forest.model";

        public static Settings settings;
        public static CustomTaskPane settingPanel = null;

        public static MarkBookHolder markBookHolder;
    }

    static class Utils
    {
        public static int ParseColumn(string col)
        {
            int temp = 0;
            foreach (var c in col)
                temp = temp * 26 + c - 'A' + 1;
            return temp;
        }

        // Line.ForeColor.RGB 和 color.TOArgb红蓝位置相反
        public static int RGBColor(Color color) => (color.B << 16) + (color.G << 8) + color.R;


        #region 获取当前实例
        public static Excel.Workbook GetActiveWorkbook()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
            if (workbook == null || workbook.FileFormat != Excel.XlFileFormat.xlOpenXMLWorkbook)
                throw new Exception("Only support OpenXML Workbook.");
            return workbook;
        }

        public static Excel.Worksheet GetActiveWorksheet(Excel.Workbook workbook)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet as Excel.Worksheet;
            if (worksheet == null)
                throw new Exception($"No Worksheet in {workbook.Name} is Active.");
            return worksheet;
        }

        public static Excel.Range GetSelectedRange()
        {
            Excel.Range range = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (range == null)
                throw new Exception("Seletion is not Excel Range.");            
            if (range.Areas.Count != 1)
                throw new Exception("Seleted Range has more than one Area.");
            return range;
        }
        #endregion
    }
}
