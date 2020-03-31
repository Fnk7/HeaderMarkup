using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace HeaderMarkup
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private static readonly string xlsx = ".xlsx";
        private static readonly string range = ".range";


        // 获取当前实例
        private static Excel.Workbook GetActiveWorkbook()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (workbook.FileFormat == Excel.XlFileFormat.xlOpenXMLWorkbook)
                return workbook;
            throw new Exception($"Only support Workbook{xlsx} file.");
        }

        private static Excel.Range GetSelectedRange()
        {
            if (!(Globals.ThisAddIn.Application.Selection is Excel.Range))
                return null;
            Excel.Range range = Globals.ThisAddIn.Application.Selection;
            return range;
        }

        // 保存
        private void buttonSaveMarkup_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = GetActiveWorkbook();
                string dataset = Share.settings.MarkupDateset;
                string markup = Share.markups.MarkupInfos(workbook);
                if (!Directory.Exists(dataset))
                    Directory.CreateDirectory(dataset);
                string name = workbook.Name;
                if (File.Exists(workbook.FullName) && name.Contains('.'))
                    name = name.Substring(0, name.LastIndexOf('.'));
                if (File.Exists(Path.Combine(dataset, name + xlsx)))
                    if (string.Equals(Path.Combine(dataset, name + xlsx), workbook.FullName, StringComparison.InvariantCultureIgnoreCase))
                        workbook.Save();
                    else if (MessageBox.Show($"Replace\t{name}{xlsx} in {dataset}？", "Replace File", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
                            == DialogResult.OK)
                    {
                        File.Delete(Path.Combine(dataset + name + xlsx));
                        File.Delete(Path.Combine(dataset + name + range));
                        workbook.SaveCopyAs(Path.Combine(dataset, name + xlsx));
                    }
                    else return;
                else
                    workbook.SaveCopyAs(Path.Combine(dataset, name + xlsx));
                using (StreamWriter streamWriter = new StreamWriter(Path.Combine(dataset, name + range)))
                    streamWriter.Write(markup);
                Share.markups.Remove(workbook);
                workbook.Close(false, Type.Missing, Type.Missing);
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message, "Failed");
            }
        }
        // 加载
        private void buttonLoadMarkup_Click(object sender, RibbonControlEventArgs e)
        {

        }


        // 标记表
        private void buttonMarkTable_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddTable(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());

        // 标记区域
        private void buttonTitleQuiteLike_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), -2);
        private void buttonTitleLittleLike_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), -1);
        private void buttonMarkHeader_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 0);
        private void buttonDataLittleLike_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 1);
        private void buttonDataQuiteLike_Click(object sender, RibbonControlEventArgs e) => Share.markups.AddMarkArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange(), 2);

        // 删除操作
        private void buttonDeleteAll_Click(object sender, RibbonControlEventArgs e) => Share.markups.DeleteAll(Globals.ThisAddIn.Application.ActiveWorkbook);
        private void buttonDeleteArea_Click(object sender, RibbonControlEventArgs e) => Share.markups.DeleteArea(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());
        private void buttonDeleteTable_Click(object sender, RibbonControlEventArgs e) => Share.markups.DeleteTable(Globals.ThisAddIn.Application.ActiveWorkbook, GetSelectedRange());


        

        // 打开或关闭设置面板
        private void buttonSettings_Click(object sender, RibbonControlEventArgs e)
        {
            if (Share.settingPanel == null)
            {
                var settingPanel = new Setting.SettingPanel();
                Share.settingPanel = Globals.ThisAddIn.CustomTaskPanes.Add(settingPanel, "Settings");
                Share.settingPanel.Visible = true;
                Share.settingPanel.Width = 400;
            }
            else if (Share.settingPanel.Visible)
                Share.settingPanel.Visible = false;
            else
                Share.settingPanel.Visible = true;
        }


        private void buttonPredict_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO
            //try
            //{
            //    var csvDataset = Share.settings.CSVDataset;
            //    if (!File.Exists(Path.Combine(csvDataset)))
            //        return;
            //    var model = Path.Combine(csvDataset, Share.modelName);
            //    HMarkupClassifier.Utils.RunPython("", $"");

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        // 训练模型
        private void buttonTrain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var csvDataset = Share.settings.CSVDataset;
                if (!Directory.Exists(csvDataset))
                    throw new Exception($"{csvDataset} doesn't exist.");
                HMarkupClassifier.Utils.RunPython("", $"train {csvDataset} {Path.Combine(csvDataset, Share.modelName)}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed");
            }
        }

        private void buttonGenerateCSV_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var dataset = Share.settings.MarkupDateset;
                var csvDataset = Share.settings.CSVDataset;
                HMarkupClassifier.Utils.ParseDataset(dataset, csvDataset);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
