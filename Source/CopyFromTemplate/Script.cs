using System;
using System.Collections.Generic;
using System.Text;

using Microsoft.Office.Interop.Excel;

using Red.Core;
using Red.Core.Logs;
using Red.Core.Office;

namespace CopyFromTemplate
{
    public class Script
    {
        public static Log Log { get; } = new Log("Script");

        public static void Execute(OfficeApps apps, Input input)
        {
            if (Flow.Interrupted)
                return;

            apps.Excel.DisplayAlerts = false;

            Log.Info("Executing Script");

            Log.Debug("Target: " + input.Workbook.FullName);
            Log.Debug("Template: " + input.Worksheet.Name);

            foreach (var name in input.NewNames)
            {
                if (Flow.Interrupted)
                    break;

                // Note: this can fail to create a unique name if interrupted
                string newName = ExcelHelper.CreateUniqueWorksheetName(input.Workbook, name);

                if (Flow.Interrupted)
                    break;

                if (newName == name)
                    Log.Debug($"Creating {name}");
                else
                    Log.Debug($"{name} exists already. Creating {newName}");

                input.Worksheet.Copy(After: input.Workbook.Sheets[apps.Excel.Sheets.Count]);

                if (Flow.Interrupted)
                    break;

                // There are issues with hidden sheets and indexing, hence using the ActiveSheet as the the target for the name
                // This relies on the fact that a copy is made active.
                Worksheet newSheet = (Worksheet) apps.Excel.ActiveSheet;
                newSheet.Name = newName;
            }

            if (!Flow.Interrupted)
            {
                Log.Debug($"Saving {input.Workbook.FullName}");
                input.Workbook.Save();
                Log.Success("Script complete");
            }
        }
    }
}
