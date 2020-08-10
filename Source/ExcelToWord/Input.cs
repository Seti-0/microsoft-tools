using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.IO;

using Microsoft.Office.Interop.Excel;
using Red.Core.Office;
using System.Linq;
using WpfToolset;
using System.ComponentModel;
using System.Net.Http.Headers;
using System.Windows.Controls;

namespace ExcelToWord
{
    public class Input
    {
        public class Source
        {
            public string Alias;
            public Workbook Workbook;
        }

        public Workbook Workbook;

        public Worksheet Template;

        public List<string> Formulae;

        public string SheetReference;

        public Dictionary<string, Source> SourcesByAlias;

        public Dictionary<string, Source> SourcesByName;
    }

    public class UserInputSource : INotifyPropertyChanged
    {
        private bool _editing;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool Editing
        {
            get => _editing;

            set
            {
                if (_editing != value)
                {
                    _editing = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Editing)));
                }
            }
        }

        // These are properties instead of fields so as to be friendly
        // to the Wpf Binding system

        public string Name { get; set; }
        public string Alias { get; set; }
        public string Path { get; set; }
    }

    public class UserInput : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public string WordFilePath;

        public string ExcelSheetNames;

        public List<UserInputSource> ExcelSources = new List<UserInputSource>();

        private void InvokeSourcesChanged()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ExcelSources)));
        }

        public void AddSource(UserInputSource source)
        {
            ExcelSources.Add(source);
            InvokeSourcesChanged();
        }

        public void RemoveSource(UserInputSource source)
        {
            ExcelSources.Remove(source);
            InvokeSourcesChanged();
        }

        public void MoveSourceUp(int index)
        {
            if (index == 0)
                return;

            if (index >= ExcelSources.Count || index < 0)
                return;

            Swap(index, index - 1);
            InvokeSourcesChanged();
        }

        public void MoveSourceDown(int index)
        {
            if (index == ExcelSources.Count - 1)
                return;

            if (index >= ExcelSources.Count || index < 0)
                return;

            Swap(index, index + 1);
            InvokeSourcesChanged();
        }

        private void Swap(int a, int b)
        {
            // Ensure that a is less than b
            if (a > b)
            {
                int c = a;
                a = b;
                b = c;
            }
            
            var A = ExcelSources[a];
            var B = ExcelSources[b];

            // The order here is important - remove the later one first
            ExcelSources.RemoveAt(b);
            ExcelSources.RemoveAt(a);

            // If the aliases are the their index, update them
            if (A.Alias == a.ToString())
                A.Alias = b.ToString();
            if (B.Alias == b.ToString())
                B.Alias = a.ToString();

            // Reinsert them at swapped positions. Again, order is important - 
            // insert the earlier one, whose place hasn't been affected, first.
            
            if (a == ExcelSources.Count)
                ExcelSources.Add(B);
            else ExcelSources.Insert(a, B);

            if (b == ExcelSources.Count)
                ExcelSources.Add(A);
            else ExcelSources.Insert(b, A);
        }

        public bool TryParse(OfficeApps apps, bool readOnly, out Input input)
        {
            input = null;
            return true;
            /*
            input = new Input();
            
            if (Flow.Interrupted)
                return false;

            if (WordFilePath == null || TemplateName == null || Formulae == null || SheetReference == null)
            {
                Script.Log.Error("Unable to parse user input - one or more fields were null");
                return false;
            }

            if (FileHelper.TryOpenWorkbook(apps, WordFilePath, readOnly, out Workbook result))
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
            */
        }
    }
}
