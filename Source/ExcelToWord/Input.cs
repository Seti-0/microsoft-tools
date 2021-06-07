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
using Red.Core.Logs;
using Microsoft.Office.Interop.Word;

namespace ExcelToWord
{
    public enum OutputFormat
    {
        PDF, Word
    }

    public class Input
    {
        public class Source
        {
            public string Alias;
            public Workbook Workbook;
        }

        public Document Template;

        public List<string> SheetNames;

        public List<Source> Sources;

        public OutputFormat OutputFormat;
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

        public string OutputFormat { get; set; }

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
            input = new Input();
            
            if (Flow.Interrupted)
                return false;

            if (WordFilePath == null || ExcelSheetNames == null || ExcelSources == null || OutputFormat == null)
            {
                Script.Log.Error("Unable to parse user input - one or more fields were null");
                return false;
            }

            input.SheetNames = StringHelper
                .Split(ExcelSheetNames)
                .ToList();

            if (input.SheetNames.Count == 0)
            {
                Script.Log.Warning("Unable to read sheet names");
                return false;
            }

            if (Flow.Interrupted)
                return false;

            if (FileHelper.TryOpenDocument(apps, WordFilePath, readOnly: true, out Document document))
                input.Template = document;

            else
            {
                // FileHelper will log the actual error
                return false;
            }

            if (Flow.Interrupted)
                return false;

            if (ExcelSources.Count == 0)
            {
                Script.Log.Warning("At least one excel source is required");
                return false;
            }

            List<Input.Source> sources = new List<Input.Source>();

            foreach (UserInputSource source in ExcelSources)
            {
                if (Flow.Interrupted)
                    return false;

                if (FileHelper.TryOpenWorkbook(apps, source.Path, true, out Workbook workbook))
                    sources.Add(new Input.Source { Alias = source.Alias, Workbook = workbook });

                // FileHelper will log the error
                else return false;
            }

            if (Flow.Interrupted)
                return false;

            if (new HashSet<string>(sources.Select(x => x.Alias)).Count < sources.Count)
            {
                Script.Log.Warning("Duplicate aliases detected");
                return false;
            }

            input.Sources = sources;

            if (Flow.Interrupted)
                return false;

            switch (OutputFormat.ToLower())
            {
                case "pdf":
                    input.OutputFormat = ExcelToWord.OutputFormat.PDF;
                    break;

                case "word document":
                    input.OutputFormat = ExcelToWord.OutputFormat.Word;
                    break;

                default:
                    Script.Log.Error("Unrecognized output format: " + OutputFormat);
                    break;
            }

            return true;
        }
    }
}
