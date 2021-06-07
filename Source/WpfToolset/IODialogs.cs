using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using System.Windows.Controls;
using Microsoft.Win32;
using Red.Core;
using Red.Core.IO;
using Red.Core.Logs;

namespace WpfToolset
{
    public class FileCheck
    {
        public Predicate<string> Predicate;

        public string SuccessMessage;

        public string AsideMessage;
    }

    public class IODialogs
    {
        /* To check filepaths after dialogs by using office interop is a nice idea, 
         * but far too slow to be feasible.
         * 
         * This could change if office is kept open for the length of the program,
         * but that brings its own pitfalls
         * 
        public static FileCheck ExcelCheck = new FileCheck
        {
            Predicate = ExcelPredicate,
            SuccessMessage = "Workbook check successful",
            AsideMessage = "Is the file accessible, and a workbook?"
        };

        public static FileCheck WordCheck = new FileCheck
        {
            Predicate = WordPredicate,
            SuccessMessage = "Document check successful",
            AsideMessage = "Is the file accessible, and a document?"
        };

        private static bool ExcelPredicate(string path)
        {
            bool result = false;

            void Check(OfficeApps apps)
            {
                result = FileHelper.TryOpenWorkbook(apps, path, readOnly: true, out _);
            }

            OfficeApps.RunExcelWithGuard(Check);

            return result;
        }

        private static bool WordPredicate(string path)
        {
            bool result = false;

            void Check(OfficeApps apps)
            {
                result = FileHelper.TryOpenDocument(apps, path, readOnly: true, out _);
            }

            OfficeApps.RunWordWithGuard(Check);

            return result;
        }
         */

        public static bool TrySelectFile(out string filePath, string title, string defaultExt)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                CheckPathExists = true,

                Multiselect = false,

                DefaultExt = defaultExt,
                DereferenceLinks = true,
                Title = title
            };

            var result = openFileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                filePath = openFileDialog.FileName;
                return true;
            }
            else
            {
                filePath = null;
                return false;
            }
        }

        public static bool TrySelectFile(TextBox textBox, Button browseButton, 
            string title, string defaultExt)
        {
            browseButton.IsEnabled = false;

            bool success = TrySelectFile(out string path, title, defaultExt);
            if (success)
                textBox.Dispatcher.Invoke(UpdateText);

            browseButton.Dispatcher.Invoke(UpdateButton);
            return success;

            /* The way Flow.cs is currently set up this causes threading issues.
             * 
             * More importantly, the process of opening office and closing
             * it fully is far to slow to be acceptable for a single check.
             * 
            void UpdateOnCheck()
            {
                bool passed = check == null || check.Predicate(path);

                if (passed)
                    textBox.Dispatcher.Invoke(UpdateText);

                browseButton.Dispatcher.Invoke(UpdateButton);

                if (check != null)
                {
                    if (passed)
                        Logs.Wpf.Debug(check.SuccessMessage);

                    else
                    {
                        Logs.Wpf.Warning("File check failed");
                        Logs.Wpf.Debug(check.AsideMessage);
                    }    
                }
            }
             */

            void UpdateText()
            {
                textBox.Clear();
                textBox.AppendText(path);

                // ScrollToEnd isn't working here, I don't know why. But scrolling to a very big number
                // does it.
                textBox.ScrollToHorizontalOffset(10000);
            }

            void UpdateButton()
            {
                browseButton.IsEnabled = true;
            }
        }
    }
}
