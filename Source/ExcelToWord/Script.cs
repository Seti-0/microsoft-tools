using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Navigation;
using Microsoft.Office.Interop.Excel;
using Red.Core;
using Red.Core.Logs;
using Red.Core.Office;
using WpfToolset;

namespace ExcelToWord
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

        private static float lastProgress = 0;

        private static void ResetProgress()
        {
            lastProgress = 0;
        }

        private static void PrintProgress(float value)
        {
            if (value - lastProgress >= 0.1)
            {
                lastProgress = value;
                Log.Debug($"{Math.Round(value * 100)}% complete");
            }
        }

        private static void CompleteProgress()
        {
            if (lastProgress != 1)
                Log.Debug($"100% complete");
        }

        public static void Execute(OfficeApps apps, Input input)
        {
            if (Flow.Interrupted)
                return;

            Log.Info("Executing Script");

            Log.Debug("Target: " + input.Workbook.FullName);
            Log.Debug("Template: " + input.Template.Name);

            Log.Debug("Checking which cells are numerical:");
            Log.PushIndent();

            var fullEnumerator = new RangeEnumerator(input.Template.UsedRange);
            bool[,] mask = new bool[fullEnumerator.Height, fullEnumerator.Width];

            ResetProgress();

            while (fullEnumerator.MoveNext())
            {
                PrintProgress(fullEnumerator.Progress);

                ExcelRange cell = fullEnumerator.Current;

                if (Flow.Interrupted)
                    break;

                bool numeric = ExcelHelper.IsCellAnywhereNumeric(apps, cell.Row, cell.Column);
                mask[fullEnumerator.RowIndex, fullEnumerator.ColumnIndex] = numeric;
            }

            CompleteProgress();
            fullEnumerator.Dispose();

            Log.PopIndent();

            if (Flow.Interrupted)
                return;

            Log.Debug("Creating summary sheets:");
            
            // Strange to double these I know, but it makes it easier to keep track of 
            // indents with interruptions this way.
            Log.PushIndent();
            Log.PushIndent();

            foreach (string formula in input.Formulae)
            {
                string name = ExcelHelper.CreateUniqueWorksheetName(input.Workbook, formula);

                Log.PopIndent();
                Log.Debug($"Creating sheet \"{name}\":");
                Log.PushIndent();

                if (Flow.Interrupted)
                    break;

                input.Template.Copy(After: input.Workbook.Sheets[apps.Excel.Sheets.Count]);

                if (Flow.Interrupted)
                    break;

                Worksheet newSheet = (Worksheet) input.Workbook.ActiveSheet;
                newSheet.Name = name;

                var maskedEnumerator = new RangeEnumerator(newSheet.UsedRange);
                maskedEnumerator.ApplyMask(mask);
                ResetProgress();

                while (maskedEnumerator.MoveNext())
                {
                    PrintProgress(maskedEnumerator.Progress);

                    ExcelRange cell = maskedEnumerator.Current;

                    if (Flow.Interrupted)
                        break;

                    string range = $"'{input.SheetReference}'!{cell.Address}";
                    string formulaText = $"={formula}({range})";

                    cell.Value = formulaText;
                }

                CompleteProgress();
                maskedEnumerator.Dispose();

                if (Flow.Interrupted)
                    break;
            }

            Log.PopIndent();
            Log.PopIndent();

            if (Flow.Interrupted)
                return;

            Log.Debug($"Saving {input.Workbook.FullName}");
            input.Workbook.Save();
            Log.Success("Script complete");
        }
    }
}
