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

        private float _tableInterval = 8f;
        private float _markInterval = 8f;
        [Category("Mark"), DisplayName("Shape Count"), ReadOnly(true)]
        public int ShapeCount { get; set; } = 0;
        [Category("Mark"), DisplayName("Interval(Table)")]
        public float TableInterval { get { return _tableInterval; } set { _tableInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Interval(Marks)")]
        public float HeaderInterval { get { return _markInterval; } set { _markInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Save Shapes")]
        public bool SaveMarkupShapes { get; set; } = false;
        [Category("Mark"), DisplayName("Save Property")]
        public bool SaveInWorkbookProperty { get; set; } = false;

        [Browsable(false)]
        public int MaxEdgeSize { get; set; } = 1000;
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
