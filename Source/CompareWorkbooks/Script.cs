using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Excel;
using Red.Core;
using Red.Core.Logs;
using Red.Core.Office;

namespace CompareWorkbooks
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

            Log.Debug("Target workbook: " + input.TargetWorkbook.FullName);
            Log.Debug("Target template: " + input.Template.Name);
            Log.Debug("Number of sources: " + input.Sources.Count);

            Log.Debug("Sources:");
            Log.PushIndent();
            foreach (var source in input.Sources)
                Log.Debug(source);
            Log.PopIndent();

            if (Flow.Interrupted)
                return;

            if (input.Template.Columns.Count < 2)
            {
                Log.Warning($"Skipping template {input.Template.Name}, since it does not seem to have a" +
                    $" second column");
            }
            else
            {
                Log.Debug("Reading template");
                Log.PushIndent();

                ExcelRange bottomRight = (ExcelRange)input.Template.UsedRange.Cells[input.Template.UsedRange.Cells.Count];
                int endRow = bottomRight.Row;
                int rowCount = endRow;

                string[] references = new string[rowCount];
                int referenceCount = 0;

                for (int i = 1; i < rowCount; i++)
                {
                    if (Flow.Interrupted)
                        break;

                    ExcelRange cell = (ExcelRange)input.Template.Cells[1 + i, 2];
                    string reference = cell?.Text?.ToString();
                    reference = reference?.Trim();

                    const string pattern = @"[A-Za-z]+\d+";

                    if (string.IsNullOrWhiteSpace(reference))
                    {
                        Log.Debug($"Skipping empty row {1 + i}");
                        continue;
                    }

                    if (!Regex.IsMatch(reference, pattern))
                    {
                        Log.Warning($"Skipping row {1 + i}: cannot parse reference \"{reference}\"");
                        continue;
                    }

                    Log.Debug($"Row {1 + i}: {reference}");

                    references[i] = reference;
                    referenceCount++;
                }

                Log.PopIndent();

                if (Flow.Interrupted)
                    return;

                if (referenceCount > 0)
                {
                    Log.Debug("Populating template.");
                    Log.PushIndent();

                    for (int i = 0; i < input.Sources.Count; i++)
                    {
                        string name = input.Sources[i];
                        Log.Debug(name);

                        ExcelRange headerCell = (ExcelRange)input.Template.Cells[1, i*2 + 3];
                        headerCell.Value = name;

                        for (int j = 0; j < references.Length; j++)
                        {
                            if (Flow.Interrupted)
                                break;

                            if (references[j] == null)
                                continue;

                            // The i + 3 is important - the columns of data should start
                            // on the third sheet column

                            if (!ExcelHelper.TrySelectWorksheet(input.TargetWorkbook,
                                out Worksheet sheetA, name, compareWords: true, verbrose: true))
                                continue;

                            if (!ExcelHelper.TrySelectWorksheet(input.OtherWorkbook,
                                out Worksheet sheetB, name, compareWords: true, verbrose: true))
                                continue;

                            string valueA = sheetA.Range[references[j]].Value2.ToString();
                            string valueB = sheetB.Range[references[j]].Value2.ToString();

                            ((ExcelRange)input.Template.Cells[j + 1, i*2 + 3]).Value = valueA;
                            ((ExcelRange)input.Template.Cells[j + 1, i*2 + 4]).Value = valueB;

                            if (Flow.Interrupted)
                                break;
                        }

                    }

                    if (Flow.Interrupted)
                        return;
                }

                Log.PopIndent();

            }

            if (Flow.Interrupted)
                return;

            Log.Debug($"Saving {input.TargetWorkbook.FullName}");
            if (input.TargetWorkbook.FullName != input.OtherWorkbook.FullName)
                Log.Debug($"(No changes made to {input.OtherWorkbook.FullName})");
            input.TargetWorkbook.Save();
            Log.Success("Script complete");
        }
    }
}
