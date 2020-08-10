using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.IO;

using Microsoft.Office.Interop.Excel;
using Red.Core.Office;
using System.Linq;

namespace CopyFromTemplate
{
    public class Input
    {
        public Workbook Workbook;

        public Worksheet Worksheet;

        public List<string> NewNames;
    }

    public class UserInput
    {
        public string FilePath { get; set; } = null;

        public string TemplateName { get; set; } = null;

        public string NewNames { get; set; } = null;

        public bool TryParse(OfficeApps apps, bool readOnly, out Input input)
        {
            input = new Input();
            
            if (Flow.Interrupted)
                return false;

            if (FilePath == null || TemplateName == null || NewNames == null)
            {
                Script.Log.Error("Unable to parse user input - one or more fields were null");
                return false;
            }

            if (FileHelper.TryOpenWorkbook(apps, FilePath, readOnly, out Workbook result))
                input.Workbook = result;

            // The FileHelper will have its own logs, no need to create new ones here
            else return false;

            if (ExcelHelper.TrySelectWorksheet(input.Workbook, out Worksheet worksheet, TemplateName, true))
            {
                input.Worksheet = worksheet;
            }
            else
            {
                input.Worksheet = null;
                Script.Log.Warning($"Unable to find sheet {TemplateName} in {input.Workbook.FullName}");
                return false;
            }

            if (Flow.Interrupted)
                return false;

            input.NewNames = StringHelper.Split(NewNames).ToList();

            if (input.NewNames.Count == 0)
            {
                Script.Log.Warning("No new sheet names given");
                return false;
            }

            return true;
        }
    }
}
