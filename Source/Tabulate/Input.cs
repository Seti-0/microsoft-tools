using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.IO;

using Microsoft.Office.Interop.Excel;
using Red.Core.Office;
using System.Linq;
using WpfToolset;

namespace Tabulate
{
    public class Input
    {
        public Workbook Workbook;

        public List<Worksheet> Templates;

        public List<Worksheet> Sources;
    }

    public class UserInput
    {
        public string FilePath { get; set; } = null;

        public string TemplateNames { get; set; } = null;

        public string SourceNames { get; set; } = null;

        public bool TryParse(OfficeApps apps, bool readOnly, out Input input)
        {
            input = new Input();

            if (Flow.Interrupted)
                return false;

            if (FilePath == null || TemplateNames == null || SourceNames == null)
            {
                Script.Log.Error("Unable to parse user input - one or more fields were null");
                return false;
            }

            IEnumerable<string> templateNames = StringHelper.Split(TemplateNames);
            IEnumerable<string> sourceNames = StringHelper.Split(SourceNames);

            if (!templateNames.Any())
            {
                Script.Log.Warning("No template names given");
                return false;
            }

            if (!sourceNames.Any())
            {
                Script.Log.Warning("No source names given");
                return false;
            }

            Workbook workbook;

            if (!FileHelper.TryOpenWorkbook(apps, FilePath, readOnly, out workbook))
                // The FileHelper will have its own logs, no need to create new ones here
                return false;

            input.Workbook = workbook;

            if (Flow.Interrupted)
                return false;

            input.Templates = new List<Worksheet>();
            foreach (string templateName in templateNames)
            {
                if (ExcelHelper.TrySelectWorksheet(workbook, out Worksheet currentSheet, templateName, compareWords: true, verbrose: true))
                    input.Templates.Add(currentSheet);
                else
                {
                    Script.Log.Warning($"Unable to find template sheet {templateName} in {workbook.FullName}");
                    return false;
                }

                if (Flow.Interrupted)
                    return false;
            }

            input.Sources = new List<Worksheet>();
            foreach (string sourceName in sourceNames)
            {
                if (ExcelHelper.TrySelectWorksheet(workbook, out Worksheet currentSheet, sourceName, compareWords: true, verbrose: true))
                    input.Sources.Add(currentSheet);
                else
                {
                    Script.Log.Warning($"Unable to find source sheet {sourceName} in {workbook.FullName}");
                    return false;
                }

                if (Flow.Interrupted)
                    return false;
            }

            return true;
        }
    }
}
