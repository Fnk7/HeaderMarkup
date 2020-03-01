namespace HeaderMarkup
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabHMarkup = this.Factory.CreateRibbonTab();
            this.groupAnnotation = this.Factory.CreateRibbonGroup();
            this.groupSave = this.Factory.CreateRibbonGroup();
            this.checkBoxDeleteShapes = this.Factory.CreateRibbonCheckBox();
            this.checkBoxSaveMarkFile = this.Factory.CreateRibbonCheckBox();
            this.checkBoxSaveMarkProperty = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonEraseMarkup = this.Factory.CreateRibbonButton();
            this.buttonRedrawMarkup = this.Factory.CreateRibbonButton();
            this.buttonShowMarkup = this.Factory.CreateRibbonButton();
            this.buttonSaveMarkup = this.Factory.CreateRibbonButton();
            this.buttonMarkTable = this.Factory.CreateRibbonButton();
            this.buttonMarkHeader = this.Factory.CreateRibbonButton();
            this.tabHMarkup.SuspendLayout();
            this.groupAnnotation.SuspendLayout();
            this.groupSave.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHMarkup
            // 
            this.tabHMarkup.Groups.Add(this.groupSave);
            this.tabHMarkup.Groups.Add(this.groupAnnotation);
            this.tabHMarkup.Label = "HMarkup";
            this.tabHMarkup.Name = "tabHMarkup";
            // 
            // groupAnnotation
            // 
            this.groupAnnotation.Items.Add(this.buttonMarkTable);
            this.groupAnnotation.Items.Add(this.buttonMarkHeader);
            this.groupAnnotation.Items.Add(this.separator2);
            this.groupAnnotation.Items.Add(this.buttonShowMarkup);
            this.groupAnnotation.Items.Add(this.buttonEraseMarkup);
            this.groupAnnotation.Items.Add(this.buttonRedrawMarkup);
            this.groupAnnotation.Label = "Annotation";
            this.groupAnnotation.Name = "groupAnnotation";
            // 
            // groupSave
            // 
            this.groupSave.Items.Add(this.checkBoxDeleteShapes);
            this.groupSave.Items.Add(this.checkBoxSaveMarkFile);
            this.groupSave.Items.Add(this.checkBoxSaveMarkProperty);
            this.groupSave.Items.Add(this.separator1);
            this.groupSave.Items.Add(this.buttonSaveMarkup);
            this.groupSave.Label = "Save";
            this.groupSave.Name = "groupSave";
            // 
            // checkBoxDeleteShapes
            // 
            this.checkBoxDeleteShapes.Checked = true;
            this.checkBoxDeleteShapes.Label = "Delete Shapes";
            this.checkBoxDeleteShapes.Name = "checkBoxDeleteShapes";
            // 
            // checkBoxSaveMarkFile
            // 
            this.checkBoxSaveMarkFile.Checked = true;
            this.checkBoxSaveMarkFile.Label = "Save MarkFile";
            this.checkBoxSaveMarkFile.Name = "checkBoxSaveMarkFile";
            // 
            // checkBoxSaveMarkProperty
            // 
            this.checkBoxSaveMarkProperty.Checked = true;
            this.checkBoxSaveMarkProperty.Label = "Save MarkProp";
            this.checkBoxSaveMarkProperty.Name = "checkBoxSaveMarkProperty";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonEraseMarkup
            // 
            this.buttonEraseMarkup.Label = "Erase Markup";
            this.buttonEraseMarkup.Name = "buttonEraseMarkup";
            this.buttonEraseMarkup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEraseMarkup_Click);
            // 
            // buttonRedrawMarkup
            // 
            this.buttonRedrawMarkup.Label = "Redraw Markup";
            this.buttonRedrawMarkup.Name = "buttonRedrawMarkup";
            this.buttonRedrawMarkup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRedrawMarkup_Click);
            // 
            // buttonShowMarkup
            // 
            this.buttonShowMarkup.Label = "Show Markup";
            this.buttonShowMarkup.Name = "buttonShowMarkup";
            this.buttonShowMarkup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowMarkup_Click);
            // 
            // buttonSaveMarkup
            // 
            this.buttonSaveMarkup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSaveMarkup.Image = global::HeaderMarkup.Properties.Resources.SaveMarkup;
            this.buttonSaveMarkup.Label = "Save Markup";
            this.buttonSaveMarkup.Name = "buttonSaveMarkup";
            this.buttonSaveMarkup.ShowImage = true;
            this.buttonSaveMarkup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSaveMarkup_Click);
            // 
            // buttonMarkTable
            // 
            this.buttonMarkTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonMarkTable.Image = global::HeaderMarkup.Properties.Resources.MarkTable;
            this.buttonMarkTable.Label = "Mark Table";
            this.buttonMarkTable.Name = "buttonMarkTable";
            this.buttonMarkTable.ShowImage = true;
            this.buttonMarkTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkTable_Click);
            // 
            // buttonMarkHeader
            // 
            this.buttonMarkHeader.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonMarkHeader.Image = global::HeaderMarkup.Properties.Resources.MarkHeader;
            this.buttonMarkHeader.Label = "Mark Header";
            this.buttonMarkHeader.Name = "buttonMarkHeader";
            this.buttonMarkHeader.ShowImage = true;
            this.buttonMarkHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkHeader_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabHMarkup);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabHMarkup.ResumeLayout(false);
            this.tabHMarkup.PerformLayout();
            this.groupAnnotation.ResumeLayout(false);
            this.groupAnnotation.PerformLayout();
            this.groupSave.ResumeLayout(false);
            this.groupSave.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnnotation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxDeleteShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSaveMarkFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSaveMarkProperty;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSaveMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEraseMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRedrawMarkup;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
