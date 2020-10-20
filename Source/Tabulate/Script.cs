using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Navigation;
using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Excel;

using Red.Core;
using Red.Core.Logs;
using Red.Core.Office;

using WpfToolset;

namespace Tabulate
{
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;

    public class Script
    {
        public static Log Log { get; } = new Log("Script");

        private struct Point2
        {
            public int X, Y;

            public Point2(int x, int y) { X = x; Y = y; }

            public override bool Equals(object obj)
            {
                if (obj is Point2 other) return other.X == X && other.Y == Y;
                else return false;
            }

            public override int GetHashCode()
            {
                // Apparently this works, no idea why.
                return X.GetHashCode() + 31 * Y.GetHashCode() + 16337;
            }
        }

        public static void Execute(OfficeApps apps, Input input)
        {
            if (Flow.Interrupted)
                return;

            Log.Info("Executing Script");

            Log.Debug("Target workbook: " + input.Workbook.FullName);
            Log.Debug("Number of templates: " + input.Templates.Count);
            Log.Debug("Number of sources: " + input.Sources.Count);

            foreach (Worksheet template in input.Templates)
            {
                if (template.Columns.Count < 2)
                {
                    Log.Warning($"Skipping template {template.Name}, since it does not seem to have a" +
                        $" second column");
                    continue;
                }

                if (Flow.Interrupted)
                    break;

                Log.Debug("Reading template: " + template.Name);
                Log.PushIndent();

                int rowCount = template.UsedRange.Rows.Count;
                string[] references = new string[rowCount];
                int referenceCount = 0;

                for (int i = 0; i < rowCount; i++)
                {
                    if (Flow.Interrupted)
                        break;

                    ExcelRange cell = (ExcelRange) template.UsedRange.Cells[1+i, 2];
                    string reference = cell?.Text?.ToString();
                    reference = reference?.Trim();

                    const string pattern = @"[A-Za-z]+\d+";

                    if (string.IsNullOrWhiteSpace(reference))
                    {
                        Log.Debug($"Skipping empty row {1+i}");
                        continue;
                    }

                    if (!Regex.IsMatch(reference, pattern))
                    {
                        Log.Warning($"Skipping row {1+i}: cannot parse reference \"{reference}\"");
                        continue;
                    }

                    Log.Debug($"Row {1+i}: {reference}");

                    references[i] = reference;
                    referenceCount++;
                }

                Log.PopIndent();

                if (Flow.Interrupted)
                    break;

                if (referenceCount == 0)
                    continue;

                Log.Debug("Populating template: " + template.Name);
                Log.PushIndent();

                for (int i = 0; i < input.Sources.Count; i++)
                {
                    string name = input.Sources[i].Name;
                    Log.Debug(name);

                    ExcelRange headerCell = (ExcelRange) template.Cells[1, i + 3];
                    headerCell.Value = name;

                    for (int j = 0; j < references.Length; j++)
                    {
                        if (Flow.Interrupted)
                            break;

                        if (references[j] == null)
                            continue;

                        string formula = $"='{name}'!{references[j]}";

                        // The i + 3 is important - the columns of data should start
                        // on the third sheet column
                        ExcelRange targetCell = (ExcelRange) template.Cells[j + 1, i + 3];
                        targetCell.Value = formula;
                    }

                    if (Flow.Interrupted)
                        break;
                }

                Log.PopIndent();

                if (Flow.Interrupted)
                    break;
            }

            if (Flow.Interrupted)
                return;

            if (Flow.Interrupted)
                return;

            Log.Debug($"Saving {input.Workbook.FullName}");
            input.Workbook.Save();
            Log.Success("Script complete");
        }
    }
}
