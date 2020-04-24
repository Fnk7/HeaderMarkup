using System;
using System.IO;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

using HeaderMarkup.DrawShape;
using HeaderMarkup.Classifiers;

namespace HeaderMarkup
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private static readonly string xlsx = ".xlsx";
        private static readonly string mark = ".mark";


        private static bool SaveMarkInfo()
        {
            var markDst = Share.settings.MarkDateset;
            if (!Directory.Exists(markDst))
            {
                MessageBox.Show("Need select a Mark Dataset Folder.");
                return false;
            }
            try
            {
                var workbook = Utils.GetActiveWorkbook();
                var baseName = workbook.Name;
                if (baseName.EndsWith(".xlsx"))
                    baseName = workbook.Name.Substring(0, workbook.Name.LastIndexOf('.'));
                string bookSavePath = Path.Combine(markDst, baseName + xlsx);
                string markSavePath = Path.Combine(markDst, baseName + mark);
                if (string.Equals(workbook.FullName, bookSavePath, StringComparison.InvariantCultureIgnoreCase))
                {
                    MessageBox.Show($"The workbook is in Mark Dataset.\nMark Dataset {markDst}");
                    return false;
                }
                if (!Share.settings.SaveMarkShapes)
                    EraseShape.EraseAll(workbook);
                workbook.SaveCopyAs(bookSavePath);
                var markInfo = Share.bookHolder.GetBook(workbook).ToString();
                using (StreamWriter markInfoWriter = new StreamWriter(markSavePath))
                    markInfoWriter.Write(markInfo);
                Share.bookHolder.Remove(workbook);
                var pathDelete = workbook.FullName;
                workbook.Close(false, Type.Missing, Type.Missing);
                if (File.Exists(pathDelete))
                    File.Delete(pathDelete);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private static void OpenNext()
        {
            if (!Share.settings.ToMarkNext)
                return;
            var files = Share.settings.FilesToMark;
            while (files.Count != 0)
            {
                try
                {
                    var file = files.Pop();
                    Globals.ThisAddIn.Application.Workbooks.Open(file);
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            Share.settings.ToMarkNext = false;
            MessageBox.Show("Finsh!");
        }

        // Save
        private void btSaveMarkInfo_Click(object sender, RibbonControlEventArgs e)
        {
            if (SaveMarkInfo())
                OpenNext();
        }

        // Drop
        private void btDeleteWorkbook_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var workbook = Utils.GetActiveWorkbook();
                var pathDelete = workbook.FullName;
                workbook.Close(false, Type.Missing, Type.Missing);
                if (File.Exists(pathDelete))
                    File.Delete(pathDelete);
                OpenNext();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Mark
        private void btMark_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var sheet = Share.bookHolder.GetSheet();
                var range = Utils.GetSelectedRange();
                // 1. Mark Table
                if (e.Control.Id == btMarkTable.Id)
                {
                    var name = sheet.AddTable(range.Address);
                    DrawTable.Draw(range, name);
                }
                // 2. Mark Others
                else
                {
                    int type = 0;
                    if (e.Control.Id == btMarkTitle.Id)
                        type = -2;
                    else if (e.Control.Id == btMarkTitleHeader.Id)
                        type = -1;
                    else if (e.Control.Id == btMarkDataHeader.Id)
                        type = 1;
                    else if (e.Control.Id == btMarkData.Id)
                        type = 2;
                    var name = sheet.AddMark(range.Address, type);
                    DrawMark.Draw(range, type, name);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Delete
        private void btDelete_Click(object sender, RibbonControlEventArgs e) 
        {
            try
            {
                if (e.Control.Id == btDeleteAll.Id)
                {
                    Share.bookHolder.GetSheet().DeletAll();
                    EraseShape.EraseAll();
                    return;
                }
                Excel.Range range = Utils.GetSelectedRange();
                if (e.Control.Id == btDeleteMark.Id)
                {
                    var name = Share.bookHolder.GetSheet().DeletMark(range.Address);
                    EraseShape.EraseByName(name);
                }else if(e.Control.Id == btDeleteTable.Id)
                {
                    var names = Share.bookHolder.GetSheet().DeletTable(range.Address);
                    EraseShape.EraseByName(names);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Settings
        private void btSettings_Click(object sender, RibbonControlEventArgs e)
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
        private void btPredict_Click(object sender, RibbonControlEventArgs e)
        {   // 展示当前Sheet的效果
            try
            {
                int clf = 0;
                if (e.Control.Id == btNB.Id)
                    clf = 1;
                else if (e.Control.Id == btNN.Id)
                    clf = 2;
                var result = Classifier.Predict(clf);
                var workbook = Utils.GetActiveWorkbook();
                var worksheet = Utils.GetActiveWorksheet(workbook);
                foreach (var (row, col) in result)
                {
                    var cell = worksheet.Cells[row, col] as Excel.Range;
                    DrawPredict.Draw(cell, $"R{row}C{col}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
