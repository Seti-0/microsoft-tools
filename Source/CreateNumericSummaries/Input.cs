﻿using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.IO;

using Microsoft.Office.Interop.Excel;
using Red.Core.Office;
using System.Linq;
using WpfToolset;

namespace CreateNumericSummaries
{
    public class Input
    {
        public Workbook Workbook;

        public Worksheet Template;

        public List<string> Formulae;

        public string SheetReference;
    }

    public class UserInput
    {
        public string FilePath { get; set; } = null;

        public string TemplateName { get; set; } = null;

        public string Formulae { get; set; } = null;

        public string SheetReference { get; set; } = null;

        public bool TryParse(OfficeApps apps, bool readOnly, out Input input)
        {
            input = new Input();
            
            if (Flow.Interrupted)
                return false;

            if (FilePath == null || TemplateName == null || Formulae == null || SheetReference == null)
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
                input.Template = worksheet;
            }
            else
            {
                input.Template = null;
                Script.Log.Warning($"Unable to find sheet {TemplateName} in {input.Workbook.FullName}");
                return false;
            }

            if (Flow.Interrupted)
                return false;

            input.Formulae = StringHelper
                // "Split" here handles empty entries and such
                .Split(Formulae)
                .ToList();

            if (input.Formulae.Count == 0)
            {
                Script.Log.Warning("No formulae given");
                return false;
            }

            var set = input.Formulae.ToHashSet();
            if (set.Count < input.Formulae.Count)
            {
                input.Formulae = set.ToList();
                Script.Log.Warning("Ignoring duplicate forumlae");
            }

            input.SheetReference = SheetReference;

            if (string.IsNullOrWhiteSpace(input.SheetReference))
            {
                Script.Log.Warning($"No sheet range given");
                return false;
            }

            if (!ExcelHelper.TryParseWorksheetRange(out string _, input.Workbook, input.SheetReference, compareWords: true))
            {
                Script.Log.Warning($"Invalid sheet reference: \"{input.SheetReference}\"");
                return false;
            }

            return true;
        }
    }
}
