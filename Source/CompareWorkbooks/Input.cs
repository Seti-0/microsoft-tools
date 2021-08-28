using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.IO;

using Microsoft.Office.Interop.Excel;
using Red.Core.Office;
using System.Linq;
using WpfToolset;

namespace CompareWorkbooks
{
    public class Input
    {
        public Workbook TargetWorkbook;

        public Workbook OtherWorkbook;

        public Worksheet Template;

        public IList<string> Sources;
    }

    public class UserInput
    {
        public string FilePathA { get; set; } = null;

        public string FilePathB { get; set; } = null;

        public string TemplateName { get; set; } = null;

        public string SourceSheetReference { get; set; } = null;

        public bool TryParse(OfficeApps apps, bool readOnly, out Input input)
        {
            input = new Input();

            if (Flow.Interrupted)
                return false;

            if (FilePathA == null || FilePathB == null || TemplateName == null || SourceSheetReference == null)
            {
                Script.Log.Error("Unable to parse user input - one or more fields were null");
                return false;
            }

            Workbook workbookA, workbookB;

            if (FileHelper.TryOpenWorkbook(apps, FilePathA, readOnly, out Workbook resultA))
                workbookA = resultA;

            // The FileHelper will have its own logs, no need to create new ones here
            else return false;

            if (FileHelper.TryOpenWorkbook(apps, FilePathB, readOnly, out Workbook resultB))
                workbookB = resultB;

            else return false;

            Worksheet worksheet;
            if (ExcelHelper.TrySelectWorksheet(workbookA, out worksheet, TemplateName, compareWords: true))
            {
                input.Template = worksheet;
                input.TargetWorkbook = workbookA;
                input.OtherWorkbook = workbookB;
            }
            else if (ExcelHelper.TrySelectWorksheet(workbookB, out worksheet, TemplateName, compareWords: true))
            {
                input.Template = worksheet;
                input.TargetWorkbook = workbookB;
                input.OtherWorkbook = workbookA;
            }
            else
            {
                input.Template = null;
                Script.Log.Warning($"Unable to find sheet {TemplateName} in {workbookA.FullName} or {workbookB.FullName}.");
                return false;
            }

            if (Flow.Interrupted)
                return false;

            if (string.IsNullOrWhiteSpace(SourceSheetReference))
            {
                Script.Log.Warning($"No source sheets given.");
                return false;
            }

            IList<Workbook> workbooks = new Workbook[] { input.TargetWorkbook, input.OtherWorkbook };
            if (!ExcelHelper.TryParseCommonWorksheetRange(out input.Sources, workbooks, SourceSheetReference))
            {
                Script.Log.Warning("Unable to read reference: '" + SourceSheetReference + "'.");
                return false;
            }

            return true;
        }
    }
}
