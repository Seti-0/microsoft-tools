using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

using Red.Core;
using WpfToolset;

namespace CompareWorkbooks
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private UserInput userInput;
        private bool selectedFileA, selectedFileB;

        public MainWindow()
        {
            InitializeComponent();

            userInput = new UserInput();

            //.IsEnabled = false;
            //SelectSheets.IsEnabled = false;

            WindowHelper.Setup(this, LogBox, Run, Cancel);

            Flow.Init();
        }

        private void SelectFileA_Click(object sender, RoutedEventArgs e)
        {
            if (IODialogs.TrySelectFile(WorkbookPathA, SelectFileA, "Select Workbook", "xlsx"))
            {
                CollectInput();
                selectedFileA = true;
                if (selectedFileB) CheckWorksheets();
            }
        }

        private void SelectFileB_Click(object sender, RoutedEventArgs e)
        {
            if (IODialogs.TrySelectFile(WorkbookPathB, SelectFileB, "Select Workbook", "xlsx"))
            {
                CollectInput();
                selectedFileB = true;
                if (selectedFileA) CheckWorksheets();
                //WindowHelper.RunWithCancel("Check for worksheets", CheckWorksheets, cancelMessage: "Cancelled worksheet check");
            }
        }

        private void SelectWorksheet_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectSharedWorksheets(TemplateName, SheetSelectionType.Range, SheetSelectionDefault.Unique);
            //ExcelDialogs.SelectWorksheet(TemplateName, WorkbookPathA.Text);
        }

        private void SelectRange_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectSharedWorksheets(SelectedSheets, SheetSelectionType.Range, SheetSelectionDefault.Shared);
            //ExcelDialogs.SelectWorksheetRange(SelectedSheets, WorkbookPathA.Text);
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
            userInput.FilePathA = WorkbookPathA.Text;
            userInput.FilePathB = WorkbookPathB.Text;
            userInput.TemplateName = TemplateName.Text;
            userInput.SourceSheetReference = SelectedSheets.Text;
        }

        private void CheckWorksheets()
        {
            //ExcelDialogs.UpdateWorksheetSelector(SelectWorksheet, TemplateName);
            //ExcelDialogs.UpdateWorksheetRangeSelector(SelectSheets, SelectedSheets);
            ExcelDialogs.CheckCommonWorksheets(new string[] { userInput.FilePathA, userInput.FilePathB });
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
