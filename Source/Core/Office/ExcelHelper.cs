using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

using Red.Core.IO;
using Red.Core.Logs;

namespace Red.Core.Office
{
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;

    public static class ExcelHelper
    {
        /// <summary>
        /// Warning-to-self: This can fail to create a unique name if the Flow Interrupted flag is set
        /// </summary>
        public static string CreateUniqueWorksheetName(Workbook workbook, string baseName)
        {
            baseName ??= "null";
            string newName = baseName;

            int i = 0;
            while (WorksheetExists(workbook, newName))
            {
                if (Flow.Interrupted)
                    break;

                newName = $"{baseName} ({++i})";
            }

            return newName;
        }

        public class Outlier
        {
            public string Name;
            public string Source;
        }

        public static void FindCommonWorksheetNames(out List<string> commonNames, out List<Outlier> outliers,
            IList<Workbook> workbooks)
        {
            commonNames = new List<string>();
            outliers = new List<Outlier>();

            List<HashSet<string>> totalSheetNames = new List<HashSet<string>>(workbooks.Count);

            foreach (var workbook in workbooks)
            {
                if (Flow.Interrupted)
                    return;

                HashSet<string> currentSheetNames = new HashSet<string>();

                foreach (Worksheet sheet in workbook.Sheets)
                {

                    currentSheetNames.Add(StringHelper.GetWords(sheet.Name));

                    if (Flow.Interrupted)
                        return;
                }

                totalSheetNames.Add(currentSheetNames);
            }

            for (int i = 0; i < workbooks.Count; i++)
            {
                if (Flow.Interrupted)
                    return;

                foreach (var name in totalSheetNames[i])
                {
                    if (commonNames.Contains(name))
                        continue;

                    if (Flow.Interrupted)
                        return;

                    bool common = true;

                    foreach (var nameCollection in totalSheetNames)
                        if (Flow.Interrupted)
                            return;
                        else if (!nameCollection.Contains(name))
                        {
                            common = false;
                            break;
                        }

                    if (common)
                        commonNames.Add(name);
                    else
                    {
                        if (TrySelectWorksheet(workbooks[i], out Worksheet worksheet, name, compareWords: true))
                            outliers.Add(new Outlier
                            {
                                Name = name,
                                Source = workbooks[i].Name
                            });
                        else
                        {
                            Log.Core.Warning("Internal Error (Recoverable)");
                            Log.Core.PushIndent();
                            Log.Core.Debug($"Error while indentifying common worksheets: name \"{name}\" extracted from but not" +
                                $" found in {workbooks[i].Name}");
                            Log.Core.Debug($"Note to self: see {nameof(ExcelHelper)}.cs for more details");
                            Log.Core.PopIndent();
                        }
                    }
                }
            }
        }

        public static bool TryParseWorksheetRange(out string parsedReference, Workbook workbook,
            string reference, bool compareWords = false, bool verbrose = false)
        {
            parsedReference = null;

            if (string.IsNullOrWhiteSpace(reference))
                return false;

            string[] elements = reference.Split(':');

            if (elements.Length == 1)
            {
                if (TrySelectWorksheet(workbook, out Worksheet worksheet, elements[0], compareWords, verbrose))
                {
                    parsedReference = worksheet.Name;
                    return true;
                }

                else return false;
            }
            else if (elements.Length == 2)
            {
                if (!TrySelectWorksheet(workbook, out Worksheet first, elements[0], compareWords, verbrose))
                    return false;

                if (!TrySelectWorksheet(workbook, out Worksheet second, elements[1], compareWords, verbrose))
                    return false;

                if (first.Index == second.Index)
                {
                    parsedReference = first.Name;
                    return true;
                }
                else if (first.Index < second.Index)
                {
                    parsedReference = $"{first.Name}:{second.Name}";
                    return true;
                }
                else
                {
                    parsedReference = $"{second.Name}:{first.Name}";
                    return true;
                }
            }

            else return false;
        }

        public static bool TryParseWorksheetRange(out IEnumerable<Worksheet> worksheets, Workbook workbook,
                    string reference, bool compareWords = false, bool verbrose = false)
        {
            worksheets = null;

            if (string.IsNullOrWhiteSpace(reference))
                return false;

            string[] elements = reference.Split(':');

            int start = 0;
            int end = 0;

            if (elements.Length == 1)
            {
                if (TrySelectWorksheet(workbook, out Worksheet worksheet, elements[0], compareWords, verbrose))
                {
                    start = worksheet.Index;
                    end = start + 1;
                }

                else return false;
            }
            else if (elements.Length == 2)
            {
                if (!TrySelectWorksheet(workbook, out Worksheet first, elements[0], compareWords, verbrose))
                    return false;

                if (!TrySelectWorksheet(workbook, out Worksheet second, elements[1], compareWords, verbrose))
                    return false;

                start = first.Index;
                end = second.Index;

                if (end < start)
                {
                    int save = start;
                    start = end;
                    end = save;
                }

                end++;
            }
            else
            {
                return false;
            }

            Worksheet[] result = new Worksheet[end - start];

            for (int i = 0; i < result.Length; i++)
            {
                result[i] = workbook.Sheets[start + i];
            }

            worksheets = result;
            return true;
        }

        public static bool TrySelectWorksheet(Workbook workbook, out Worksheet worksheet,
            string name, bool compareWords = false, bool verbrose = false)
        {
            worksheet = null;

            if (Flow.Interrupted)
                return false;

            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == name || (compareWords && StringHelper.CompareWords(sheet.Name, name)))
                {
                    worksheet = sheet;
                    return true;
                }

                if (Flow.Interrupted)
                    break;
            }

            if (verbrose) Log.Core.Debug($"Unable to find worksheet \"{name}\" in workbook \"{workbook.Name}\"");
            return false;
        }
        public static bool WorksheetExists(Workbook workbook, string name,
            bool compareWords = false)
        {
            return TrySelectWorksheet(workbook, out var _, name, compareWords);
        }

        public static IEnumerable<string> GetSheetNames(OfficeApps apps, string workbookPath)
        {
            if (FileHelper.TryOpenWorkbook(apps, workbookPath, true, out Workbook result))
            {
                return GetSheetNames(result);
            }

            else return new string[0];
        }

        public static IEnumerable<string> GetSheetNames(Workbook workbook)
        {
            foreach (Worksheet sheet in workbook.Sheets)
                yield return sheet.Name;
        }

        public static bool IsCellAnywhereNumeric(OfficeApps apps, int row, int column)
        {
            foreach (Worksheet sheet in apps.Excel.Sheets)
            {
                ExcelRange cell = sheet.Cells[row, column];
                string content = cell.Value?.ToString() ?? "";

                if (double.TryParse(content, out var _))
                    return true;
            }

            return false;
        }
    }
}
