﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using System.Windows;
using System.Windows.Controls;
using Red.Core;
using Red.Core.IO;
using Red.Core.Office;
using WpfToolset;

namespace ExcelToWord
{
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private UserInput userInput;

        public MainWindow()
        {
            InitializeComponent();

            userInput = new UserInput();
            ExcelSources.ItemsSource = userInput.ExcelSources;

            SaveAsType.SelectedIndex = 0;

            WindowHelper.Setup(this, LogBox, Run, Cancel);

            Flow.Init();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            IODialogs.TrySelectFile(DocumentPath, "Select Workbook", "xlsx");
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            CollectInput();
            WindowHelper.RunWithCancel("Run Script", RunGuardedScript, cancelMessage: "Script cancelled");
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Flow.Interrupt(Flow.Reason.Cancel);
        }

        private void AddSource_Click(object sender, RoutedEventArgs e)
        {
            if (IODialogs.TrySelectFile(out string path, "Select Excel Source", ".xlsx"))
            {
                string bookName = Path.GetFileNameWithoutExtension(path);

                userInput.AddSource(new UserInputSource()
                {
                    Alias = userInput.ExcelSources.Count.ToString(),
                    Name = bookName,
                    Path = path
                });
                ExcelSources.Items.Refresh();

                ExcelDialogs.CheckCommonWorksheets(userInput.ExcelSources.Select(x => x.Path));
            }
        }

        private void RemoveSource_Click(object sender, RoutedEventArgs e)
        {
            var selectedIndex = ExcelSources.SelectedIndex;

            var source = ExcelSources.SelectedItem as UserInputSource;
            if (source != null)
                userInput.RemoveSource(source);
            ExcelSources.Items.Refresh();

            if (selectedIndex >= ExcelSources.Items.Count)
                selectedIndex--;

            ExcelSources.SelectedIndex = selectedIndex;

            ExcelDialogs.CheckCommonWorksheets(userInput.ExcelSources.Select(x => x.Path));
        }

        private void MoveSourceUp_Click(object sender, RoutedEventArgs e)
        {
            userInput.MoveSourceUp(ExcelSources.SelectedIndex);
            ExcelSources.Items.Refresh();
        }

        private void MoveSourceDown_Click(object sender, RoutedEventArgs e)
        {
            userInput.MoveSourceDown(ExcelSources.SelectedIndex);
            ExcelSources.Items.Refresh();
        }

        private void CollectInput()
        {
            userInput.WordFilePath = DocumentPath.Text;
        }

        private void RunGuardedScript()
        {
            OfficeApps.RunExcelWithGuard(RunScript);
        }

        private void RunScript(OfficeApps apps)
        {
            if (Flow.Interrupted)
                return;

            if (userInput.TryParse(apps, readOnly: false, out var input))
                Script.Execute(apps, input);
        }

        private void TextBlock_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                var textblock = sender as FrameworkElement;
                var userInput = textblock.DataContext as UserInputSource;

                userInput.Editing = true;
            }
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var control = sender as FrameworkElement;
            var userInput = control.DataContext as UserInputSource;

            userInput.Editing = false;
        }
    }
}
