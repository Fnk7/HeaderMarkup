using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Windows.Forms;

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
            if (context == null || context.Instance == null || provider == null)
                return value;
            try
            {
                string dir = (string)value;
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (Directory.Exists(dir))
                    dialog.SelectedPath = dir;
                else
                    dialog.SelectedPath = "D:\\";
                using (dialog)
                    if(dialog.ShowDialog() == DialogResult.OK)
                        dir = dialog.SelectedPath;
                return dir;
            }
            catch (Exception)
            {
                return value;
            }
        }
    }
}
