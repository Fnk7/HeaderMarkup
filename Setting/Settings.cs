using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Design;

using System.IO;

namespace HeaderMarkup.Setting
{
    public class Settings
    {
        [Category("Path"), DisplayName("CSV Dataset"),
            Editor(typeof(DirectorySelector), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string CSVDataset { get; set; } = Share.defualtPath;

        [Category("Path"), DisplayName("Mark Dataset"),
            Editor(typeof(DirectorySelector), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string MarkDateset { get; set; } = Share.defualtPath;

        [Category("Path"), DisplayName("To Mark Dataset"),
            Editor(typeof(DirectorySelector), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string ToMarkDateset { get; set; } = Share.defualtPath;

        [Category("Path"), DisplayName("To Mark Next")]
        public bool ToMarkNext
        {
            get
            {
                if (FilesToMark == null)
                    return false;
                return true;
            }
            set 
            {
                if (value && Directory.Exists(ToMarkDateset))
                {
                    FilesToMark = new Stack<string>(Directory.GetFiles(ToMarkDateset, "*.xlsx"));
                    if (FilesToMark.Count == 0)
                        FilesToMark = null;
                }
                else
                    FilesToMark = null;
            }
        }

        [Category("Path"), DisplayName("To Mark Files"), ReadOnly(true)]
        public Stack<string> FilesToMark { get; set; } = null;

        private string pythonFiles = string.Empty;
        [Category("Classifier"), DisplayName("Python Files"),
            Editor(typeof(DirectorySelector), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string PythonFiles
        {
            get
            {
                return pythonFiles;
            }
            set
            {
                if (File.Exists(Path.Combine(value, "main.py")))
                    pythonFiles = value;
            }
        }

        private string classifier = string.Empty;
        [Category("Classifier"), DisplayName("Classifier"), TypeConverter(typeof(ClassifierItems))]
        public string Classifier
        {
            get
            {
                if (!Directory.Exists(pythonFiles))
                    return string.Empty;
                if (!File.Exists(Path.Combine(pythonFiles, classifier)))
                    return string.Empty;
                return classifier;
            }
            set
            {
                classifier = value;
            }
        }

        [Category("Classifier"), DisplayName("Temp Folder"),
            Editor(typeof(DirectorySelector), typeof(UITypeEditor)),
            TypeConverter(typeof(TypeConverter))]
        public string TempDir { get; set; } = Path.GetTempPath();


        private float tableInterval = 8f;
        private float markInterval = 8f;
        private float tableLineWeight = 1.5f;
        private float markLineWeight = 1.5f;
        [Category("Mark"), DisplayName("Interval (Table)")]
        public float TableInterval { get { return tableInterval; } set { tableInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Interval (Mark)")]
        public float MarkInterval { get { return markInterval; } set { markInterval = Math.Max(4f, Math.Min(value, 16f)); } }
        [Category("Mark"), DisplayName("Line Wieght (Table)")]
        public float TableLineWeight { get { return tableLineWeight; } set { tableLineWeight = Math.Max(1f, Math.Min(value, 2f)); } }
        [Category("Mark"), DisplayName("Line Wieght (Mark)")]
        public float MarkLineWeight { get { return markLineWeight; } set { markLineWeight = Math.Max(1f, Math.Min(value, 2f)); } }
        [Category("Mark"), DisplayName("Save Mark Shapes")]
        public bool SaveMarkShapes { get; set; } = false;

        [Browsable(false)]
        public string TableShapeName { get; set; } = "TableShape:";
        [Browsable(false)]
        public string TableEdgeName { get; set; } = "TableEdge:";
        [Browsable(false)]
        public string MarkShapeName { get; set; } = "MarkShape:";
        [Browsable(false)]
        public string MarkLineName { get; set; } = "MarkLine:";
        [Browsable(false)]
        public string PredictShapeName { get; set; } = "PredictShape:";
        [Browsable(false)]
        public string PredictLineName { get; set; } = "PredictLine:";
    }
}
