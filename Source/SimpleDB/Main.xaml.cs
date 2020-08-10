using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

using System.Windows;
using Red.Core;

using WpfToolset;

namespace SimpleDB
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private UserInput userInput;

        public MainWindow()
        {
            InitializeComponent();

            userInput = new UserInput();

            SelectWorksheet.IsEnabled = false;
            SelectRange.IsEnabled = false;

            WindowHelper.Setup(this, LogBox, Run, Cancel);

            Flow.Init();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            if (IODialogs.TrySelectFile(WorkbookPath, "Select Workbook", "xlsx"))
            {
                CollectInput();
                WindowHelper.RunWithCancel("Check for worksheets", CheckWorksheets, cancelMessage: "Cancelled worksheet check");
            }
        }

        private void SelectWorksheet_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectWorksheet(TemplateName, WorkbookPath.Text);
        }

        private void SelectRange_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectWorksheetRange(SheetRange, WorkbookPath.Text);
        }

        private void Run_Click(object sender, RoutedEventArgs e)
        {
            CollectInput();
            WindowHelper.RunWithCancel("Run script", RunGuardedScript, cancelMessage: "Script cancelled");
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Flow.Interrupt(Flow.Reason.Cancel);
        }

        private void CollectInput()
        {
            userInput.FilePath = WorkbookPath.Text;
            userInput.TemplateName = TemplateName.Text;
            userInput.SheetReference = SheetRange.Text;
        }

        private void CheckWorksheets()
        {
            ExcelDialogs.CheckSheets(userInput.FilePath);
            ExcelDialogs.UpdateWorksheetSelector(SelectWorksheet, TemplateName);
            ExcelDialogs.UpdateWorksheetRangeSelector(SelectRange, SheetRange);
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
    }
}
