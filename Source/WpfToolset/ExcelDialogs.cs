using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Windows.Controls;

using Red.Core;
using Red.Core.Office;
using Red.Core.IO;

namespace WpfToolset
{
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    public enum SheetSelectionType
    {
        List, Range
    }

    public enum SheetSelectionDefault
    {
        Shared, Unique
    }

    public static class ExcelDialogs
    {
        private class SharedWorksheet
        {
            public string SharedName;
            public bool IsCommon;
            public string SourceWorksheet;
        }

        private static IList<SharedWorksheet> sharedWorksheets;

        private static IEnumerable<string> worksheetNames;

        public static bool SheetRangeDialogAvailable => worksheetNames != null && worksheetNames.Any();

        public static void CheckCommonWorksheets(IEnumerable<string> workbookPaths)
        {
            if (workbookPaths.Count() < 2)
                sharedWorksheets = new List<SharedWorksheet>(0);

            void Update(OfficeApps apps)
            {
                UpdateSharedWorksheets(apps, workbookPaths);
            }

            void safeUpdate() => OfficeApps.RunExcelWithGuard(Update);
            WindowHelper.RunWithCancel("Find Common Worksheets", safeUpdate, "Cancelled checking common sheets");
        }

        private static void UpdateSharedWorksheets(OfficeApps apps, IEnumerable<string> paths)
        {
            var workbooks = OpenWorkbooks(apps, paths).ToList();

            if (workbooks.Count == 0)
                sharedWorksheets = null;

            sharedWorksheets = new List<SharedWorksheet>();

            ExcelHelper.FindCommonWorksheetNames(out var commonNames, out var outliers, workbooks);

            if (Flow.Interrupted)
                return;

            foreach (var common in commonNames)
            {
                if (Flow.Interrupted)
                    return;

                sharedWorksheets.Add(new SharedWorksheet
                {
                    IsCommon = true,
                    SharedName = common,
                    SourceWorksheet = ""
                });
            }

            foreach (var outlier in outliers)
            {
                if (Flow.Interrupted)
                    return;

                sharedWorksheets.Add(new SharedWorksheet
                {
                    IsCommon = false,
                    SharedName = outlier.Name,
                    SourceWorksheet = outlier.Source
                });
            }
        }

        private static IEnumerable<Workbook> OpenWorkbooks(OfficeApps apps, IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                if (FileHelper.TryOpenWorkbook(apps, path, readOnly: true, out var workbook))
                    yield return workbook;
            }
        }

        public static void CheckSheets(string workbookPath)
        {
            OfficeApps.RunExcelWithGuard((apps) => UnsafeCheckSheets(apps, workbookPath));
        }

        public static void UpdateWorksheetSelector(Button button, TextBox text)
        {
            bool success = worksheetNames != null && worksheetNames.Any();

            button.Dispatcher.Invoke(
                () => button.IsEnabled = success);

            if (success)
            {
                text.Dispatcher.Invoke(
                    () => text.Text = worksheetNames.First());
            }
        }

        public static void UpdateWorksheetRangeSelector(Button button, TextBox text)
        {
            if (button != null)
            {
                button.Dispatcher.Invoke(
                    () => button.IsEnabled = SheetRangeDialogAvailable);
            }

            if (text != null && SheetRangeDialogAvailable)
            {
                // This is silly and convoluted I know. 
                // I just want the first, second and last elements

                var enumerator = worksheetNames.GetEnumerator();
                enumerator.MoveNext();

                string first = enumerator.Current;
                string second = null;
                string last = first;

                if (enumerator.MoveNext())
                {
                    second = enumerator.Current;
                    last = second;

                    while (enumerator.MoveNext())
                        last = enumerator.Current;
                }

                string range = $"{second ?? first}:{last}";

                text.Dispatcher.Invoke(
                    () => text.Text = range);
            }    
        }

        private static void UnsafeCheckSheets(OfficeApps apps, string workbookPath)
        {
            worksheetNames = ExcelHelper.GetSheetNames(apps, workbookPath)
                .ToList(); // It is important that the generator be evaluated now, and not 
            //later when the excel app is likely closed or doing something else.

            if (Flow.Interrupted)
                return;

            bool success = worksheetNames.Any();

            if (success)
                Logs.Wpf.Info($"Found {worksheetNames.Count()} sheet(s)");

            else
            {
                Logs.Wpf.Warning("No sheets found");
            }

        }

        public static bool SelectSharedWorksheets(TextBox textBox, 
            SheetSelectionType selectionType = SheetSelectionType.List,
            SheetSelectionDefault selectionDefault = SheetSelectionDefault.Shared)
        {
            if (sharedWorksheets != null)
            {
                SelectWorksheetDialog worksheetDialog = new SelectWorksheetDialog(null, SelectionMode.Multiple);

                List<SelectWorksheetDialog.Item> list = new List<SelectWorksheetDialog.Item>();
                foreach (var item in sharedWorksheets)
                {
                    string name = item.SharedName;

                    if (!item.IsCommon)
                        name += $" ({item.SourceWorksheet})";

                    list.Add(new SelectWorksheetDialog.Item { 
                        UserData = item.SharedName,
                        Content = name, 
                        Primary = item.IsCommon == (selectionDefault == SheetSelectionDefault.Shared) 
                    });
                }

                worksheetDialog.Items = list;
                worksheetDialog.SelectAll(primaryOnly: true);

                bool? success = worksheetDialog.ShowDialog();

                if (success.HasValue && success.Value && worksheetDialog.Results != null)
                {
                    string result;
                    if (selectionType == SheetSelectionType.Range)
                    {
                        result = worksheetDialog.Results.First();

                        if (worksheetDialog.Results.Count > 1)
                            result += ":" + worksheetDialog.Results.Last();
                    }
                    else
                    {
                        result = string.Join(',', worksheetDialog.Results);
                    }

                    textBox.Text = result;
                    textBox.ScrollToEnd();
                    return true;
                }
            }

            return false;
        }

        public static bool SelectWorksheet(TextBox textBox, string workbookPath)
        {
            if (worksheetNames != null)
            {
                var worksheetDialog = new SelectWorksheetDialog(workbookPath, SelectionMode.Single);
                worksheetDialog.ItemContent = worksheetNames;

                bool? success = worksheetDialog.ShowDialog();

                if (success.HasValue && success.Value && worksheetDialog.Result != null)
                {
                    textBox.Text = worksheetDialog.Result;
                    textBox.ScrollToEnd();
                    return true;
                }
            }

            return false;
        }

        public static bool SelectWorksheets(TextBox textBox, string workbookPath)
        {
            if (worksheetNames != null)
            {
                var worksheetDialog = new SelectWorksheetDialog(workbookPath, SelectionMode.Extended);

                worksheetDialog.ItemContent = worksheetNames;

                bool? success = worksheetDialog.ShowDialog();

                if (success.HasValue && success.Value && worksheetDialog.Results != null)
                {
                    textBox.Text = string.Join(',', worksheetDialog.Results);
                    textBox.ScrollToEnd();
                    return true;
                }
            }

            return false;
        }

        public static bool SelectWorksheetRange(TextBox textBox, string workbookPath)
        {
            if (worksheetNames != null)
            {
                var worksheetDialog = new SelectWorksheetDialog(workbookPath, SelectionMode.Extended);
                
                worksheetDialog.ItemContent = worksheetNames;

                if (worksheetNames.Count() > 1)
                    worksheetDialog.SetSelectedRange(1, worksheetNames.Count());

                bool? success = worksheetDialog.ShowDialog();

                if (success.HasValue && success.Value && worksheetDialog.Results != null)
                {
                    textBox.Text = worksheetDialog.Results.First();

                    if (worksheetDialog.Results.Count > 1)
                        textBox.Text +=  ":" + worksheetDialog.Results.Last();

                    textBox.ScrollToEnd();
                    return true;
                }
            }

            return false;
        }
    }
}
