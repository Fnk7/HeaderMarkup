using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HeaderMarkup.Classifiers
{
    class Classifier
    {
        public static readonly string[] clf_names = { "random_forest.pkl", "naive_bayes.pkl", "neural_net.pkl" };

        public static List<(int, int)> Predict(int clf)
        {
            var tempdir = Path.GetTempPath(); // TODO
            var workbook = Utils.GetActiveWorkbook();
            var worksheet = Utils.GetActiveWorksheet(workbook);
            var tempxlsx = Path.Combine(tempdir, "temp_sheet.xlsx");
            var tempInput = Path.Combine(tempdir, "temp_sheet.csv");
            var tempOutput = Path.Combine(tempdir, "temp_result.csv");
            workbook.SaveCopyAs(tempxlsx);
            HMarkupClassifier.Tools.ParseWorksheet(tempxlsx, worksheet.Name, tempInput);
            string[] arguments = new string[3];
            arguments[0] = clf_names[clf];
            arguments[1] = tempInput;
            arguments[2] = tempOutput;
            // TODO
            var (code, output, error) = Python.RunPython("D:\\Workspace\\Python\\HeaderClf\\main.py", arguments);
            if (code != 0)
                throw new Exception(error);
            return ReadResult(tempOutput);
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
