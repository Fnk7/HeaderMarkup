using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

using HeaderMarkup.Markup;

namespace HeaderMarkup
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private static readonly string xlsx = ".xlsx";
        private static readonly string range = ".range";

        // 保存 TODO
        private void buttonSaveMarkup_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = Utils.GetActiveWorkbook();
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
                MessageBox.Show(ex.Message);
            }
        }
        // 加载
        private void buttonLoadMarkup_Click(object sender, RibbonControlEventArgs e)
        {
            // TODO
            MessageBox.Show("TODO");
        }


        // 标记Table
        private void buttonMarkTable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = Utils.GetActiveWorkbook();
                Excel.Range range = Utils.GetSelectedRange();
                Share.markups.AddTable(workbook, range);
                MarkSheet sheet = Share.markBookHolder.GetMarkSheet();
                sheet.AddTable(range.Address);
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
        // 标记Header
        private void buttonMarkHeader_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int type = 0;
                if (e.Control.Id == buttonTitleQuiteLike.Id)
                    type = -2;
                else if (e.Control.Id == buttonTitleLittleLike.Id)
                    type = -1;
                else if (e.Control.Id == buttonDataLittleLike.Id)
                    type = 1;
                else if (e.Control.Id == buttonDataQuiteLike.Id)
                    type = 2;
                Excel.Workbook workbook = Utils.GetActiveWorkbook();
                Excel.Range range = Utils.GetSelectedRange();
                Share.markups.AddMarkArea(workbook, range, type);
                MarkSheet sheet = Share.markBookHolder.GetMarkSheet();
                sheet.AddHeader(range.Address, type);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // 删除操作
        private void buttonDelete_Click(object sender, RibbonControlEventArgs e) 
        {
            try
            {
                Excel.Workbook workbook = Utils.GetActiveWorkbook();
                if (e.Control.Id == buttonDeleteAll.Id)
                {
                    Share.markups.DeleteAll(workbook);
                    Share.markBookHolder.GetMarkSheet().DeletAll();
                    return;
                }
                Excel.Range range = Utils.GetSelectedRange();
                if (e.Control.Id == buttonDeleteArea.Id)
                {
                    Share.markups.DeleteArea(workbook, range);
                    Share.markBookHolder.GetMarkSheet().DeletHeader(range.Address);
                }else if(e.Control.Id == buttonDeleteTable.Id)
                {
                    Share.markups.DeleteTable(workbook, range);
                    Share.markBookHolder.GetMarkSheet().DeletTable(range.Address);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // 打开,关闭设置面板
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
            try
            {
                MarkSheet sheet = Share.markBookHolder.GetMarkSheet();
                MessageBox.Show(sheet.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 训练模型
        private void buttonTrain_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MarkBook book = Share.markBookHolder.GetMarkBook();
                MessageBox.Show(book.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonGenerateCSV_Click(object sender, RibbonControlEventArgs e)
        {
            //try
            //{
            //    var dataset = Share.settings.MarkupDateset;
            //    var csvDataset = Share.settings.CSVDataset;
            //    HMarkupClassifier.Utils.ParseDataset(dataset, csvDataset);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
    }
}
