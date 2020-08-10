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

namespace $projectname$
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

            Log.Debug($"Saving {input.Workbook.FullName}");
            input.Workbook.Save();
            Log.Success("Script complete");
        }
    }
}
