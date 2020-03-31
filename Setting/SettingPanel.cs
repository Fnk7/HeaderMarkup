using System.Windows.Forms;

namespace HeaderMarkup.Setting
{
    public partial class SettingPanel : UserControl
    {
        public SettingPanel()
        {
            InitializeComponent();
            Settings.SelectedObject = Share.settings;
        }
    }
}
