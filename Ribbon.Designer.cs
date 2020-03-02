﻿namespace HeaderMarkup
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
            this.groupSave = this.Factory.CreateRibbonGroup();
            this.checkBoxSaveShapes = this.Factory.CreateRibbonCheckBox();
            this.checkBoxSaveMarkFile = this.Factory.CreateRibbonCheckBox();
            this.checkBoxSaveMarkProperty = this.Factory.CreateRibbonCheckBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.buttonSaveToDataset = this.Factory.CreateRibbonButton();
            this.groupAnnotation = this.Factory.CreateRibbonGroup();
            this.buttonMarkTable = this.Factory.CreateRibbonButton();
            this.buttonMarkHeader = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.buttonEraseShapes = this.Factory.CreateRibbonButton();
            this.buttonRedrawShapes = this.Factory.CreateRibbonButton();
            this.buttonReset = this.Factory.CreateRibbonButton();
            this.tabHMarkup.SuspendLayout();
            this.groupSave.SuspendLayout();
            this.groupAnnotation.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabHMarkup
            // 
            this.tabHMarkup.Groups.Add(this.groupSave);
            this.tabHMarkup.Groups.Add(this.groupAnnotation);
            this.tabHMarkup.Label = "HMarkup";
            this.tabHMarkup.Name = "tabHMarkup";
            this.tabHMarkup.Position = this.Factory.RibbonPosition.AfterOfficeId("TabInsert");
            // 
            // groupSave
            // 
            this.groupSave.Items.Add(this.checkBoxSaveShapes);
            this.groupSave.Items.Add(this.checkBoxSaveMarkFile);
            this.groupSave.Items.Add(this.checkBoxSaveMarkProperty);
            this.groupSave.Items.Add(this.separator1);
            this.groupSave.Items.Add(this.buttonSaveToDataset);
            this.groupSave.Label = "Save";
            this.groupSave.Name = "groupSave";
            // 
            // checkBoxSaveShapes
            // 
            this.checkBoxSaveShapes.Label = "Save Shapes";
            this.checkBoxSaveShapes.Name = "checkBoxSaveShapes";
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
            // buttonSaveToDataset
            // 
            this.buttonSaveToDataset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSaveToDataset.Image = global::HeaderMarkup.Properties.Resources.SaveMarkup;
            this.buttonSaveToDataset.Label = "Save to Dataset";
            this.buttonSaveToDataset.Name = "buttonSaveToDataset";
            this.buttonSaveToDataset.ShowImage = true;
            this.buttonSaveToDataset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSaveToDataset_Click);
            // 
            // groupAnnotation
            // 
            this.groupAnnotation.Items.Add(this.buttonMarkTable);
            this.groupAnnotation.Items.Add(this.buttonMarkHeader);
            this.groupAnnotation.Items.Add(this.separator2);
            this.groupAnnotation.Items.Add(this.buttonEraseShapes);
            this.groupAnnotation.Items.Add(this.buttonRedrawShapes);
            this.groupAnnotation.Items.Add(this.buttonReset);
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
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // buttonEraseShapes
            // 
            this.buttonEraseShapes.Image = global::HeaderMarkup.Properties.Resources.Erase;
            this.buttonEraseShapes.Label = "Erase";
            this.buttonEraseShapes.Name = "buttonEraseShapes";
            this.buttonEraseShapes.ShowImage = true;
            this.buttonEraseShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEraseShapes_Click);
            // 
            // buttonRedrawShapes
            // 
            this.buttonRedrawShapes.Image = global::HeaderMarkup.Properties.Resources.Redraw;
            this.buttonRedrawShapes.Label = "Redraw";
            this.buttonRedrawShapes.Name = "buttonRedrawShapes";
            this.buttonRedrawShapes.ShowImage = true;
            this.buttonRedrawShapes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRedrawShapes_Click);
            // 
            // buttonReset
            // 
            this.buttonReset.Image = global::HeaderMarkup.Properties.Resources.Reset;
            this.buttonReset.Label = "Reset";
            this.buttonReset.Name = "buttonReset";
            this.buttonReset.ShowImage = true;
            this.buttonReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonReset_Click);
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabHMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnnotation;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMarkHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSaveShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSaveMarkFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSaveMarkProperty;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSaveToDataset;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEraseShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRedrawShapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonReset;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
