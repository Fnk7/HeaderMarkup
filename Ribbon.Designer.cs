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
            this.grpSaveOrDrop = this.Factory.CreateRibbonGroup();
            this.btSaveMarkInfo = this.Factory.CreateRibbonButton();
            this.btDeleteWorkbook = this.Factory.CreateRibbonButton();
            this.grpAnnotate = this.Factory.CreateRibbonGroup();
            this.btMarkTable = this.Factory.CreateRibbonButton();
            this.btMarkHeader = this.Factory.CreateRibbonButton();
            this.btMarkData = this.Factory.CreateRibbonButton();
            this.btMarkTitle = this.Factory.CreateRibbonButton();
            this.labelNot = this.Factory.CreateRibbonLabel();
            this.btMarkDataHeader = this.Factory.CreateRibbonButton();
            this.btMarkTitleHeader = this.Factory.CreateRibbonButton();
            this.labelSeem = this.Factory.CreateRibbonLabel();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btDeleteAll = this.Factory.CreateRibbonButton();
            this.btDeleteMark = this.Factory.CreateRibbonButton();
            this.btDeleteTable = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btTODO1 = this.Factory.CreateRibbonButton();
            this.btTODO2 = this.Factory.CreateRibbonButton();
            this.btTODO3 = this.Factory.CreateRibbonButton();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.btSettings = this.Factory.CreateRibbonButton();
            this.tabHMarkup.SuspendLayout();
            this.grpSaveOrDrop.SuspendLayout();
            this.grpAnnotate.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHMarkup
            // 
            this.tabHMarkup.Groups.Add(this.grpSaveOrDrop);
            this.tabHMarkup.Groups.Add(this.grpAnnotate);
            this.tabHMarkup.Groups.Add(this.group1);
            this.tabHMarkup.Groups.Add(this.groupSettings);
            this.tabHMarkup.Label = "HMarkup";
            this.tabHMarkup.Name = "tabHMarkup";
            this.tabHMarkup.Position = this.Factory.RibbonPosition.AfterOfficeId("TabInsert");
            // 
            // grpSaveOrDrop
            // 
            this.grpSaveOrDrop.Items.Add(this.btSaveMarkInfo);
            this.grpSaveOrDrop.Items.Add(this.btDeleteWorkbook);
            this.grpSaveOrDrop.Label = "Save/Drop";
            this.grpSaveOrDrop.Name = "grpSaveOrDrop";
            // 
            // btSaveMarkInfo
            // 
            this.btSaveMarkInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btSaveMarkInfo.Image = global::HeaderMarkup.Properties.Resources.save;
            this.btSaveMarkInfo.Label = "Save MarkInfo";
            this.btSaveMarkInfo.Name = "btSaveMarkInfo";
            this.btSaveMarkInfo.ShowImage = true;
            this.btSaveMarkInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btSaveMarkInfo_Click);
            // 
            // btDeleteWorkbook
            // 
            this.btDeleteWorkbook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btDeleteWorkbook.Image = global::HeaderMarkup.Properties.Resources.delete;
            this.btDeleteWorkbook.Label = "Delete Workbook";
            this.btDeleteWorkbook.Name = "btDeleteWorkbook";
            this.btDeleteWorkbook.ShowImage = true;
            this.btDeleteWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btDeleteWorkbook_Click);
            // 
            // grpAnnotate
            // 
            this.grpAnnotate.Items.Add(this.btMarkTable);
            this.grpAnnotate.Items.Add(this.btMarkHeader);
            this.grpAnnotate.Items.Add(this.btMarkData);
            this.grpAnnotate.Items.Add(this.btMarkTitle);
            this.grpAnnotate.Items.Add(this.labelNot);
            this.grpAnnotate.Items.Add(this.btMarkDataHeader);
            this.grpAnnotate.Items.Add(this.btMarkTitleHeader);
            this.grpAnnotate.Items.Add(this.labelSeem);
            this.grpAnnotate.Items.Add(this.separator2);
            this.grpAnnotate.Items.Add(this.btDeleteAll);
            this.grpAnnotate.Items.Add(this.btDeleteMark);
            this.grpAnnotate.Items.Add(this.btDeleteTable);
            this.grpAnnotate.Label = "Annotate";
            this.grpAnnotate.Name = "grpAnnotate";
            // 
            // btMarkTable
            // 
            this.btMarkTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btMarkTable.Image = global::HeaderMarkup.Properties.Resources.table;
            this.btMarkTable.Label = "Mark Table";
            this.btMarkTable.Name = "btMarkTable";
            this.btMarkTable.ShowImage = true;
            this.btMarkTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // btMarkHeader
            // 
            this.btMarkHeader.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btMarkHeader.Image = global::HeaderMarkup.Properties.Resources.header;
            this.btMarkHeader.Label = "Mark Header";
            this.btMarkHeader.Name = "btMarkHeader";
            this.btMarkHeader.ShowImage = true;
            this.btMarkHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // btMarkData
            // 
            this.btMarkData.Image = global::HeaderMarkup.Properties.Resources.data;
            this.btMarkData.Label = "Data";
            this.btMarkData.Name = "btMarkData";
            this.btMarkData.ShowImage = true;
            this.btMarkData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // btMarkTitle
            // 
            this.btMarkTitle.Image = global::HeaderMarkup.Properties.Resources.title;
            this.btMarkTitle.Label = "Title";
            this.btMarkTitle.Name = "btMarkTitle";
            this.btMarkTitle.ShowImage = true;
            this.btMarkTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // labelNot
            // 
            this.labelNot.Label = "Not Header";
            this.labelNot.Name = "labelNot";
            // 
            // btMarkDataHeader
            // 
            this.btMarkDataHeader.Image = global::HeaderMarkup.Properties.Resources.dataHeader;
            this.btMarkDataHeader.Label = "D-Header";
            this.btMarkDataHeader.Name = "btMarkDataHeader";
            this.btMarkDataHeader.ShowImage = true;
            this.btMarkDataHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // btMarkTitleHeader
            // 
            this.btMarkTitleHeader.Image = global::HeaderMarkup.Properties.Resources.titleHeader;
            this.btMarkTitleHeader.Label = "T-Header";
            this.btMarkTitleHeader.Name = "btMarkTitleHeader";
            this.btMarkTitleHeader.ShowImage = true;
            this.btMarkTitleHeader.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btMark_Click);
            // 
            // labelSeem
            // 
            this.labelSeem.Label = "Seem Header";
            this.labelSeem.Name = "labelSeem";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btDeleteAll
            // 
            this.btDeleteAll.Image = global::HeaderMarkup.Properties.Resources.deleteShape;
            this.btDeleteAll.Label = "Delete All";
            this.btDeleteAll.Name = "btDeleteAll";
            this.btDeleteAll.ShowImage = true;
            this.btDeleteAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btDelete_Click);
            // 
            // btDeleteMark
            // 
            this.btDeleteMark.Image = global::HeaderMarkup.Properties.Resources.deleteShape;
            this.btDeleteMark.Label = "Delete Mark";
            this.btDeleteMark.Name = "btDeleteMark";
            this.btDeleteMark.ShowImage = true;
            this.btDeleteMark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btDelete_Click);
            // 
            // btDeleteTable
            // 
            this.btDeleteTable.Image = global::HeaderMarkup.Properties.Resources.deleteShape;
            this.btDeleteTable.Label = "Delete Table";
            this.btDeleteTable.Name = "btDeleteTable";
            this.btDeleteTable.ShowImage = true;
            this.btDeleteTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btDelete_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btTODO1);
            this.group1.Items.Add(this.btTODO2);
            this.group1.Items.Add(this.btTODO3);
            this.group1.Label = "Predict";
            this.group1.Name = "group1";
            // 
            // btTODO1
            // 
            this.btTODO1.Image = global::HeaderMarkup.Properties.Resources.run;
            this.btTODO1.Label = "Generate CSV";
            this.btTODO1.Name = "btTODO1";
            this.btTODO1.ShowImage = true;
            this.btTODO1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerateCSV_Click);
            // 
            // btTODO2
            // 
            this.btTODO2.Image = global::HeaderMarkup.Properties.Resources.run;
            this.btTODO2.Label = "Train Model";
            this.btTODO2.Name = "btTODO2";
            this.btTODO2.ShowImage = true;
            this.btTODO2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTrain_Click);
            // 
            // btTODO3
            // 
            this.btTODO3.Image = global::HeaderMarkup.Properties.Resources.run;
            this.btTODO3.Label = "Predict Header";
            this.btTODO3.Name = "btTODO3";
            this.btTODO3.ShowImage = true;
            this.btTODO3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btPredict_Click);
            // 
            // groupSettings
            // 
            this.groupSettings.Items.Add(this.btSettings);
            this.groupSettings.Label = "Settings";
            this.groupSettings.Name = "groupSettings";
            // 
            // btSettings
            // 
            this.btSettings.Image = global::HeaderMarkup.Properties.Resources.settings;
            this.btSettings.Label = "Settings";
            this.btSettings.Name = "btSettings";
            this.btSettings.ShowImage = true;
            this.btSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btSettings_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabHMarkup);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabHMarkup.ResumeLayout(false);
            this.tabHMarkup.PerformLayout();
            this.grpSaveOrDrop.ResumeLayout(false);
            this.grpSaveOrDrop.PerformLayout();
            this.grpAnnotate.ResumeLayout(false);
            this.grpAnnotate.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAnnotate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSaveOrDrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btSaveMarkInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkDataHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btMarkTitleHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btDeleteAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btDeleteMark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btDeleteTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btDeleteWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btTODO2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btTODO3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btTODO1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelSeem;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelNot;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
