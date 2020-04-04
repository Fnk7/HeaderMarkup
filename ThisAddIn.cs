using HeaderMarkup.Setting;
using HeaderMarkup.Markup;

namespace HeaderMarkup
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Share.settings = new Settings();
            Share.bookHolder = new BookHolder();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Share.bookHolder = null;
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
