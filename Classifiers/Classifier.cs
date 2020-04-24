using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HeaderMarkup.Classifiers
{
    class Classifier
    {
        public static readonly string[] clf_names = { "random_forest.pkl", "naive_bayes.pkl", "neural_net.pkl" };

        public static List<(int, int)> Predict()
        {
            var tempdir = Share.settings.TempDir;
            string[] arguments = new string[3];
            arguments[0] = Share.settings.Classifier;
            if (arguments[0] == string.Empty)
                throw new Exception("Need Select a Classifier.");
            arguments[1] = Path.Combine(tempdir, "temp_sheet.csv");
            arguments[2] = Path.Combine(tempdir, "temp_result.csv");
            var tempxlsx = Path.Combine(tempdir, "temp_sheet.xlsx");
            var workbook = Utils.GetActiveWorkbook();
            var worksheet = Utils.GetActiveWorksheet(workbook);
            workbook.SaveCopyAs(tempxlsx);
            HMarkupClassifier.Tools.ParseWorksheet(tempxlsx, worksheet.Name, arguments[1]);
            var pythonMain = Path.Combine(Share.settings.PythonFiles, "main.py");
            var (code, output, error) = Python.RunPython(pythonMain, arguments);
            if (code != 0)
                throw new Exception(error);
            return ReadResult(arguments[2]);
        }


        public static List<(int, int)> ReadResult(string resultFile)
        {
            Dictionary<(int, int), int> result = new Dictionary<(int, int), int>();
            try
            {
                using (StreamReader reader = new StreamReader(resultFile))
                {
                    var title = reader.ReadLine();
                    if (title != "type,row,col")
                        throw new Exception("Title Error.");
                    while (reader.ReadLine() is string line)
                    {
                        var values = line.Split(',');
                        int rst = Convert.ToInt32(values[0]);
                        int row = Convert.ToInt32(values[1]);
                        int col = Convert.ToInt32(values[2]);
                        result[(row, col)] = rst;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error Occur in Result File.", ex);
            }
            return (from cell in result where cell.Value == 1 select cell.Key).ToList();
        }

    }
}
