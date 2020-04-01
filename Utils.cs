using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace HeaderMarkup
{
    static class Utils
    {
        public static int ParseColumn(string col)
        {
            int temp = 0;
            foreach (var c in col)
                temp = temp * 26 + c - 'A' + 1;
            return temp;
        }

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
