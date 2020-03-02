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
                if (!Directory.Exists(Properties.Settings.Default.MarkupDatasetAnnotatedPath)) 
                    Directory.CreateDirectory(Properties.Settings.Default.MarkupDatasetAnnotatedPath);
                if (string.Equals(Properties.Settings.Default.MarkupDatasetAnnotatedPath.TrimEnd(Path.DirectorySeparatorChar), workbook.Path.TrimEnd(Path.DirectorySeparatorChar), StringComparison.InvariantCultureIgnoreCase) && workbook.FileFormat == Excel.XlFileFormat.xlOpenXMLWorkbook)
                {
                    MessageBox.Show("不能保存" + Properties.Settings.Default.MarkupDatasetAnnotatedPath + "中的文件！");
                    return;
                }
                string name = workbook.Name;
                if(!string.IsNullOrEmpty(workbook.Path) && name.Contains('.'))
                    name = Path.GetFileNameWithoutExtension(workbook.FullName);
                if (File.Exists(Path.Combine(Properties.Settings.Default.MarkupDatasetAnnotatedPath + name + ".xlsx")))
                {
                    int i = 1;
                    while (File.Exists(Path.Combine(Properties.Settings.Default.MarkupDatasetAnnotatedPath, name + " (" + i.ToString() + ").xlsx")))
                        i++;
                    DialogResult dialogResult = MessageBox.Show("在" + Properties.Settings.Default.MarkupDatasetAnnotatedPath + 
                        "中有同名文件。\n\tYes:删除同名文件。\n\tNO:使用 " + name + " (" + i.ToString() + ").xlsx",
                        "文件重名", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        File.Delete(Path.Combine(Properties.Settings.Default.MarkupDatasetAnnotatedPath + name + ".xlsx"));
                        File.Delete(Path.Combine(Properties.Settings.Default.MarkupDatasetAnnotatedPath + name + ".rg"));
                    }
                    else if (dialogResult == DialogResult.No)
                        name += " (" + i.ToString() + ")";
                    else return;
                }
                MessageBox.Show("Hello");
                string markup = Markups.markups.SaveMarkup(workbook, checkBoxSaveShapes.Checked, checkBoxSaveMarkProperty.Checked);
                workbook.SaveAs(Filename: Properties.Settings.Default.MarkupDatasetAnnotatedPath + name + ".xlsx", FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
                using (StreamWriter streamWriter = new StreamWriter(Path.Combine(Properties.Settings.Default.MarkupDatasetAnnotatedPath, name + ".rg")))
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

        private void buttonMarkTable_Click(object sender, RibbonControlEventArgs e) => Markups.markups.MarkTable(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());

        private void buttonMarkHeader_Click(object sender, RibbonControlEventArgs e) => Markups.markups.MarkHeader(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());

        private void buttonEraseShapes_Click(object sender, RibbonControlEventArgs e) => Markups.markups.EraseShapes(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

        private void buttonRedrawShapes_Click(object sender, RibbonControlEventArgs e) => Markups.markups.RedrawShapes(Globals.ThisAddIn.Application.ActiveWorkbook);

        private void buttonReset_Click(object sender, RibbonControlEventArgs e) => Markups.markups.Reset(Globals.ThisAddIn.Application.ActiveWorkbook);
    }
}
