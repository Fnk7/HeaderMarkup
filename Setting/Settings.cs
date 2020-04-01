using System;
using System.ComponentModel;
using System.Drawing.Design;

namespace HeaderMarkup.Setting
{
    public class Settings
    {
        [Category("Path"), DisplayName("Markup Dataset"), 
            Editor(typeof(DirectorySelectTypeEditor), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string MarkupDateset { get; set; } = Share.defualtDataset;

        [Category("Path"), DisplayName("CSV Dataset"), 
            Editor(typeof(DirectorySelectTypeEditor), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string CSVDataset { get; set; } = Share.defualtCSV;

        private float tableInterval = 8f;
        private float headerInterval = 8f;
        private float tableLineWeight = 1.5f;
        private float headerLineWeight = 1.5f;
        [Category("Mark"), DisplayName("Interval (Table)")]
        public float TableInterval { get { return tableInterval; } set { tableInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Interval (Header)")]
        public float HeaderInterval { get { return headerInterval; } set { headerInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Line Wieght (Table)")]
        public float TableLineWeight { get { return tableLineWeight; } set { tableLineWeight = Math.Max(1f, Math.Min(value, 2f)); } }
        [Category("Mark"), DisplayName("Line Wieght (Header)")]
        public float HeaderLineWeight { get { return headerLineWeight; } set { headerLineWeight = Math.Max(1f, Math.Min(value, 2f)); } }
        [Category("Mark"), DisplayName("Save Mark Shapes")]
        public bool SaveMarkShapes { get; set; } = false;

        [Browsable(false)]
        public string TableShapeName { get; set; } = "MarkupTableShape:";
        [Browsable(false)]
        public string TableEdgeName { get; set; } = "MarkupTableEdge:";
        [Browsable(false)]
        public string HeaderShapeName { get; set; } = "MarkupHeaderShape:";
        [Browsable(false)]
        public string HeaderLineName { get; set; } = "MarkupHeaderLine:";
    }
}
