using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Navigation;

using Microsoft.Office.Interop.Excel;
using Red.Core;
using Red.Core.Logs;
using Red.Core.Office;
using WpfToolset;

namespace SimpleDB
{
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;

    public class Script
    {
        public static Log Log { get; } = new Log("Script");

        public static void Execute(OfficeApps apps, Input input)
        {
            if (Flow.Interrupted)
                return;

            Log.Info("Executing Script");

            input.Template.Copy(After: input.Workbook.Sheets[apps.Excel.Sheets.Count]);

            if (Flow.Interrupted)
                return;
            
            Worksheet active = (Worksheet) apps.Excel.ActiveSheet;
            active.Name = ExcelHelper.CreateUniqueWorksheetName(input.Workbook, "Summary");

            if (Flow.Interrupted)
                return;

            ExcelHelper.TryParseWorksheetRange(out IEnumerable<Worksheet> worksheets, input.Workbook, input.SheetReference,
                                                compareWords: true, verbrose: true);

            if (Flow.Interrupted)
                return;

            RangeEnumerator enumerator = new RangeEnumerator(active.UsedRange);

            Progress.Init(Log.Debug);
            Progress.Reset();

            while (enumerator.MoveNext())
            {
                if (Flow.Interrupted)
                    return;

                ExcelRange cell = enumerator.Current;
                string text = cell.Value?.ToString();
                if (text == null) continue;

                string command = text.Trim().ToLower();

                string referencePattern = "\\$\\s*([\\d\\w]+)\\s*\\$";

                if (command == "$sheetname$")
                {
                    int i = 0;
                    foreach (Worksheet sheet in worksheets)
                    {
                        if (Flow.Interrupted)
                            return;

                        // Note the self: .Cells[i,j] is 1-based, not 0-based!
                        ExcelRange recordCell = (ExcelRange) active.Cells[1 + enumerator.RowIndex, 1 + enumerator.ColumnIndex + i];
                        recordCell.Value = sheet.Name;
                        i++;

                    }
                }
                else
                {
                    Match match = Regex.Match(command.ToUpper(), referencePattern);
                    if (match.Success)
                    {
                        string reference = match.Groups[1].Value;

                        try
                        {
                            int i = 0;
                            foreach (Worksheet sheet in worksheets)
                            {
                                if (Flow.Interrupted)
                                    return;

                                ExcelRange sourceCell = (ExcelRange)sheet.Range[reference];
                                ExcelRange destinationCell = (ExcelRange)active.Cells[1 + enumerator.RowIndex, 1 + enumerator.ColumnIndex + i];

                                destinationCell.Value = sourceCell.Value?.ToString();
                                i++;
                            }
                        }
                        catch(Exception e)
                        {
                            Log.Warning($"Failed to apply reference {reference}");
                            Log.Debug("(Is it a valid reference?)");
                            
                            if (AppHelper.DebugFlag)
                                Log.Debug("Note to self:", e);
                        }
                    }
                }

                Progress.Report(enumerator.Progress);
            }

            if (Flow.Interrupted)
                return;

            Progress.Complete();

            Log.Debug($"Saving {input.Workbook.FullName}");
            input.Workbook.Save();
            Log.Success("Script complete");
        }
    }
}
