using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HeaderMarkup.Classifiers
{
    static class Python
    {
        public static (int, string, string) RunPython(string pythonFile, string[] argument, bool background = true)
        {
            try
            {
                using (System.Diagnostics.Process process = new System.Diagnostics.Process())
                {
                    if (background)
                        process.StartInfo.FileName = "pythonw";
                    else
                        process.StartInfo.FileName = "python";
                    if (argument == null || argument.Length == 0)
                        process.StartInfo.Arguments = $"{pythonFile}";
                    else
                        process.StartInfo.Arguments = $"{pythonFile} {string.Join(" ", argument)}";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    process.Start();
                    process.WaitForExit();
                    var output = process.StandardOutput.ReadToEnd();
                    var error = process.StandardError.ReadToEnd();
                    return (process.ExitCode, output, error);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Need Python Environment", ex);
            }
        }
    }
}
