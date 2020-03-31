using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace HeaderMarkup.Setting
{
    class DirectorySelectTypeEditor : UITypeEditor
    {
        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            if (context == null || context.Instance == null)
                return base.GetEditStyle(context);
            return UITypeEditorEditStyle.Modal;
        }

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            IWindowsFormsEditorService editorService;
            if (context == null || context.Instance == null || provider == null)
                return Share.defualtDataset;
            try
            {
                editorService = (IWindowsFormsEditorService)provider.GetService(typeof(IWindowsFormsEditorService));
                string directory = (string)value;
                try
                {
                    if (Path.GetFullPath(directory) != directory)
                        directory = Share.defualtDataset;
                }
                catch (Exception)
                {
                    directory = Share.defualtDataset;
                }
                FolderBrowserDialog folder = new FolderBrowserDialog();
                folder.SelectedPath = directory;
                folder.Description = "请选择一个文件夹";
                using (folder)
                    if(folder.ShowDialog() == DialogResult.OK)
                        directory = folder.SelectedPath;
                return directory;
            }
            finally
            {
                editorService = null;
            }
        }
    }
}
