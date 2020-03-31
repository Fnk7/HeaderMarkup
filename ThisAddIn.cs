using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools;

using HeaderMarkup.Setting;

namespace HeaderMarkup
{
    static class Share
    {
        public static readonly string defualtDataset = "D:\\Temp\\Markup";
        public static readonly string defualtCSV = "D:\\Temp\\CSV";
        public static readonly string modelName = "forest.model";

        public static Settings settings;
        public static Markups markups;
        public static CustomTaskPane settingPanel = null;
    }


    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Share.settings = new Settings();
            Share.markups = new Markups();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Share.markups = null;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
