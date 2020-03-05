using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Windows.Forms;
using System.IO;

namespace HeaderMarkup
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonSaveToDataset_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                string markup = Markups.markups.SaveMarkup(workbook, checkBoxSaveShapes.Checked, checkBoxSaveMarkProperty.Checked);
                string annotatedPath = Properties.Settings.Default.DatasetAnnotatedPath, name = workbook.Name, xlsx = ".xlsx", range = ".range";
                if (!Directory.Exists(annotatedPath))
                    Directory.CreateDirectory(annotatedPath);
                if (!string.IsNullOrEmpty(workbook.Path) && name.Contains('.'))
                    name = Path.GetFileNameWithoutExtension(workbook.FullName);
                if (File.Exists(Path.Combine(annotatedPath, name + xlsx)))
                    if (string.Equals(Path.Combine(annotatedPath, name + xlsx), workbook.FullName, StringComparison.InvariantCultureIgnoreCase))
                        workbook.Save();
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("已有" + name + xlsx + "标注文件，是否替换？", "替换文件", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.OK)
                        {
                            File.Delete(Path.Combine(annotatedPath + name + xlsx));
                            File.Delete(Path.Combine(annotatedPath + name + range));
                        }
                        else return;
                        workbook.SaveAs(Filename: annotatedPath + name + xlsx, FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                else
                    workbook.SaveAs(Filename: annotatedPath + name + xlsx, FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
                using (StreamWriter streamWriter = new StreamWriter(Path.Combine(annotatedPath, name + range)))
                    streamWriter.Write(markup);
                Markups.markups.Remove(workbook);
                workbook.Close();
            }
            catch (Exception)
            {
            }
        }

        private Excel.Range GetSelectedRange()
        {
            if (!(Globals.ThisAddIn.Application.Selection is Excel.Range))
                return null;
            Excel.Range range = Globals.ThisAddIn.Application.Selection;
            return range;
        }

        private void buttonMarkTable_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddTable(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());

        private void buttonEraseShapes_Click(object sender, RibbonControlEventArgs e) => Markups.markups.EraseShapes(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

        private void buttonRedrawShapes_Click(object sender, RibbonControlEventArgs e) => Markups.markups.RedrawShapes(Globals.ThisAddIn.Application.ActiveWorkbook);

        private void buttonReset_Click(object sender, RibbonControlEventArgs e) => Markups.markups.Reset(Globals.ThisAddIn.Application.ActiveWorkbook);

        // 标记区域
        private void buttonTitleQuiteLike_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), -2);
        private void buttonTitleLittleLike_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), -1);
        private void buttonMarkHeader_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 0);
        private void buttonDataLittleLike_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 1);
        private void buttonDataQuiteLike_Click(object sender, RibbonControlEventArgs e) => Markups.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 2);
    }
}
