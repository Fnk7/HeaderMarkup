using System;
using System.Linq;
using System.IO;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

using HeaderMarkup.Markup;
using HeaderMarkup.DrawShape;

namespace HeaderMarkup
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private static readonly string xlsx = ".xlsx";
        private static readonly string mark = ".mark";

        // 保存Markup
        private void buttonSaveMarkup_Click(object sender, RibbonControlEventArgs e)
        {
            var dataset = Share.settings.MarkupDateset;
            try
            {
                Excel.Workbook workbook = Utils.GetActiveWorkbook();
                var markup = Share.markBookHolder.GetMarkBook(workbook).ToString();
                if (!Directory.Exists(dataset))
                    Directory.CreateDirectory(dataset);
                string bookName = workbook.Name;
                if (File.Exists(workbook.FullName) && bookName.Contains('.'))
                    bookName = bookName.Substring(0, bookName.LastIndexOf('.'));
                string bookSavePath = Path.Combine(dataset, bookName + xlsx);
                string markSavePath = Path.Combine(dataset, bookName + mark);
                if (!Share.settings.SaveMarkShapes)
                    EraseShape.EraseAll(workbook);
                if (!File.Exists(bookSavePath))
                    workbook.SaveCopyAs(bookSavePath);
                else if (string.Equals(workbook.FullName, bookSavePath, StringComparison.InvariantCultureIgnoreCase))
                    workbook.Save();
                else if (MessageBox.Show($"Replace\t{bookSavePath}？", "Replace File",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    workbook.SaveCopyAs(bookSavePath);
                else return;
                using (StreamWriter markWriter = new StreamWriter(markSavePath))
                    markWriter.Write(markup);
                Share.markBookHolder.Remove(workbook);
                workbook.Close(false, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 加载Markup
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
                var name = Share.markBookHolder.GetMarkSheet().AddTable(range.Address);
                DrawTable.Draw(range, name);
            }
            catch(Exception ex)
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
                Excel.Range range = Utils.GetSelectedRange();
                MarkSheet sheet = Share.markBookHolder.GetMarkSheet();
                string name = sheet.AddHeader(range.Address, type);
                DrawHeader.Draw(range, type, name);
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
                if (e.Control.Id == buttonDeleteAll.Id)
                {
                    Share.markBookHolder.GetMarkSheet().DeletAll();
                    EraseShape.EraseAll();
                    return;
                }
                Excel.Range range = Utils.GetSelectedRange();
                if (e.Control.Id == buttonDeleteArea.Id)
                {
                    var name = Share.markBookHolder.GetMarkSheet().DeletHeader(range.Address);
                    EraseShape.EraseByName(name);
                }else if(e.Control.Id == buttonDeleteTable.Id)
                {
                    var names = Share.markBookHolder.GetMarkSheet().DeletTable(range.Address);
                    EraseShape.EraseByName(names);
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

        // TODO
        private void buttonPredict_Click(object sender, RibbonControlEventArgs e)
        {   // 展示当前MarkSheet的效果
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

        // 训练模型 TODO
        private void buttonTrain_Click(object sender, RibbonControlEventArgs e)
        {   // 展示当前MarkBook的效果
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

        // TODO
        private void buttonGenerateCSV_Click(object sender, RibbonControlEventArgs e)
        {
        }
    }
}
