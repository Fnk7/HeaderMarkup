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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.tabHMarkup = this.Factory.CreateRibbonTab();
            this.groupSave = this.Factory.CreateRibbonGroup();
            this.buttonSaveMarkup = this.Factory.CreateRibbonButton();
            this.buttonLoadMarkup = this.Factory.CreateRibbonButton();
            this.groupAnnotation = this.Factory.CreateRibbonGroup();
            this.buttonMarkTable = this.Factory.CreateRibbonButton();
            this.buttonMarkHeader = this.Factory.CreateRibbonButton();
            this.buttonTitleQuiteLike = this.Factory.CreateRibbonButton();
            this.buttonTitleLittleLike = this.Factory.CreateRibbonButton();
            this.labelTitleLike = this.Factory.CreateRibbonLabel();
            this.buttonDataQuiteLike = this.Factory.CreateRibbonButton();
            this.buttonDataLittleLike = this.Factory.CreateRibbonButton();
            this.labelDataLike = this.Factory.CreateRibbonLabel();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonDeleteAll = this.Factory.CreateRibbonButton();
            this.buttonDeleteArea = this.Factory.CreateRibbonButton();
            this.buttonDeleteTable = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonGenerateCSV = this.Factory.CreateRibbonButton();
            this.buttonTrain = this.Factory.CreateRibbonButton();
            this.buttonPredict = this.Factory.CreateRibbonButton();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.buttonSettings = this.Factory.CreateRibbonButton();
            this.tabHMarkup.SuspendLayout();
            this.groupSave.SuspendLayout();
            this.groupAnnotation.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHMarkup
            // 
            this.tabHMarkup.Groups.Add(this.groupSave);
            this.tabHMarkup.Groups.Add(this.groupAnnotation);
            this.tabHMarkup.Groups.Add(this.group1);
            this.tabHMarkup.Groups.Add(this.groupSettings);
            this.tabHMarkup.Label = "HMarkup";
            this.tabHMarkup.Name = "tabHMarkup";
            this.tabHMarkup.Position = this.Factory.RibbonPosition.AfterOfficeId("TabInsert");
            // 
            // groupSave
            // 
            this.groupSave.Items.Add(this.buttonSaveMarkup);
            this.groupSave.Items.Add(this.buttonLoadMarkup);
            this.groupSave.Label = "Save Load";
            this.groupSave.Name = "groupSave";
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
            // buttonLoadMarkup
            // 
            this.buttonLoadMarkup.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonLoadMarkup.Image = global::HeaderMarkup.Properties.Resources.SaveMarkup;
            this.buttonLoadMarkup.Label = "Load Markup";
            this.buttonLoadMarkup.Name = "buttonLoadMarkup";
            this.buttonLoadMarkup.ShowImage = true;
            this.buttonLoadMarkup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLoadMarkup_Click);
            // 
            // groupAnnotation
            // 
            this.groupAnnotation.Items.Add(this.buttonMarkTable);
            this.groupAnnotation.Items.Add(this.buttonMarkHeader);
            this.groupAnnotation.Items.Add(this.buttonTitleQuiteLike);
            this.groupAnnotation.Items.Add(this.buttonTitleLittleLike);
            this.groupAnnotation.Items.Add(this.labelTitleLike);
            this.groupAnnotation.Items.Add(this.buttonDataQuiteLike);
            this.groupAnnotation.Items.Add(this.buttonDataLittleLike);
            this.groupAnnotation.Items.Add(this.labelDataLike);
            this.groupAnnotation.Items.Add(this.separator2);
            this.groupAnnotation.Items.Add(this.buttonDeleteAll);
            this.groupAnnotation.Items.Add(this.buttonDeleteArea);
            this.groupAnnotation.Items.Add(this.buttonDeleteTable);
            this.groupAnnotation.Label = "Annotation";
            this.groupAnnotation.Name = "groupAnnotation";
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
            // buttonTitleQuiteLike
            // 
            this.buttonTitleQuiteLike.Image = global::HeaderMarkup.Properties.Resources.Quite;
            this.buttonTitleQuiteLike.Label = "Quite";
            this.buttonTitleQuiteLike.Name = "buttonTitleQuiteLike";
            this.buttonTitleQuiteLike.ShowImage = true;
            this.buttonTitleQuiteLike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkHeader_Click);
            // 
            // buttonTitleLittleLike
            // 
            this.buttonTitleLittleLike.Image = global::HeaderMarkup.Properties.Resources.Little;
            this.buttonTitleLittleLike.Label = "Little";
            this.buttonTitleLittleLike.Name = "buttonTitleLittleLike";
            this.buttonTitleLittleLike.ShowImage = true;
            this.buttonTitleLittleLike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkHeader_Click);
            // 
            // labelTitleLike
            // 
            this.labelTitleLike.Label = "Title Like";
            this.labelTitleLike.Name = "labelTitleLike";
            // 
            // buttonDataQuiteLike
            // 
            this.buttonDataQuiteLike.Image = global::HeaderMarkup.Properties.Resources.Quite;
            this.buttonDataQuiteLike.Label = "Quite";
            this.buttonDataQuiteLike.Name = "buttonDataQuiteLike";
            this.buttonDataQuiteLike.ShowImage = true;
            this.buttonDataQuiteLike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkHeader_Click);
            // 
            // buttonDataLittleLike
            // 
            this.buttonDataLittleLike.Image = global::HeaderMarkup.Properties.Resources.Little;
            this.buttonDataLittleLike.Label = "Little";
            this.buttonDataLittleLike.Name = "buttonDataLittleLike";
            this.buttonDataLittleLike.ShowImage = true;
            this.buttonDataLittleLike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMarkHeader_Click);
            // 
            // labelDataLike
            // 
            this.labelDataLike.Label = "Data Like";
            this.labelDataLike.Name = "labelDataLike";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonDeleteAll
            // 
            this.buttonDeleteAll.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteAll.Image")));
            this.buttonDeleteAll.Label = "Delete All";
            this.buttonDeleteAll.Name = "buttonDeleteAll";
            this.buttonDeleteAll.ShowImage = true;
            this.buttonDeleteAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDelete_Click);
            // 
            // buttonDeleteArea
            // 
            this.buttonDeleteArea.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteArea.Image")));
            this.buttonDeleteArea.Label = "Delete Area";
            this.buttonDeleteArea.Name = "buttonDeleteArea";
            this.buttonDeleteArea.ShowImage = true;
            this.buttonDeleteArea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDelete_Click);
            // 
            // buttonDeleteTable
            // 
            this.buttonDeleteTable.Image = ((System.Drawing.Image)(resources.GetObject("buttonDeleteTable.Image")));
            this.buttonDeleteTable.Label = "Delete Table";
            this.buttonDeleteTable.Name = "buttonDeleteTable";
            this.buttonDeleteTable.ShowImage = true;
            this.buttonDeleteTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDelete_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonGenerateCSV);
            this.group1.Items.Add(this.buttonTrain);
            this.group1.Items.Add(this.buttonPredict);
            this.group1.Label = "Predict";
            this.group1.Name = "group1";
            // 
            // buttonGenerateCSV
            // 
            this.buttonGenerateCSV.Image = ((System.Drawing.Image)(resources.GetObject("buttonGenerateCSV.Image")));
            this.buttonGenerateCSV.Label = "Generate CSV";
            this.buttonGenerateCSV.Name = "buttonGenerateCSV";
            this.buttonGenerateCSV.ShowImage = true;
            this.buttonGenerateCSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerateCSV_Click);
            // 
            // buttonTrain
            // 
            this.buttonTrain.Image = global::HeaderMarkup.Properties.Resources.Erase;
            this.buttonTrain.Label = "Train Model";
            this.buttonTrain.Name = "buttonTrain";
            this.buttonTrain.ShowImage = true;
            this.buttonTrain.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonTrain_Click);
            // 
            // buttonPredict
            // 
            this.buttonPredict.Image = global::HeaderMarkup.Properties.Resources.Little;
            this.buttonPredict.Label = "Predict Header";
            this.buttonPredict.Name = "buttonPredict";
            this.buttonPredict.ShowImage = true;
            this.buttonPredict.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPredict_Click);
            // 
            // groupSettings
            // 
            this.groupSettings.Items.Add(this.buttonSettings);
            this.groupSettings.Label = "Settings";
            this.groupSettings.Name = "groupSettings";
            // 
            // buttonSettings
            // 
            this.buttonSettings.Image = global::HeaderMarkup.Properties.Resources.Little;
            this.buttonSettings.Label = "Settings";
            this.buttonSettings.Name = "buttonSettings";
            this.buttonSettings.ShowImage = true;
            this.buttonSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSettings_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabHMarkup);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabHMarkup.ResumeLayout(false);
            this.tabHMarkup.PerformLayout();
            this.groupSave.ResumeLayout(false);
            this.groupSave.PerformLayout();
            this.groupAnnotation.ResumeLayout(false);
            this.groupAnnotation.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnnotation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSaveMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTitleQuiteLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTitleLittleLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelTitleLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDataQuiteLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDataLittleLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelDataLike;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteArea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeleteTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLoadMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonTrain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPredict;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenerateCSV;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
