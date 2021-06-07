using Red.Core;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WpfToolset;

namespace Tabulate
{
    /// <summary>
    /// Interaction logic for Main.xaml
    /// </summary>
    public partial class Main : Window
    {
        private UserInput userInput;

        public Main()
        {
            InitializeComponent();

            userInput = new UserInput();

            SelectTemplates.IsEnabled = false;
            SelectSources.IsEnabled = false;

            WindowHelper.Setup(this, LogBox, Run, Cancel);

            Flow.Init();
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            if (IODialogs.TrySelectFile(WorkbookPath, SelectFile, "Select Workbook", "xlsx"))
            {
                CollectInput();
                WindowHelper.RunWithCancel("Check worksheets", CheckWorksheets, cancelMessage: "Cancelled worksheet check");
            }
        }

        private void SelectTemplates_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectWorksheets(TemplateRange, WorkbookPath.Text);
        }

        private void SelectSources_Click(object sender, RoutedEventArgs e)
        {
            ExcelDialogs.SelectWorksheets(SourceRange, WorkbookPath.Text);
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

        private void CollectInput()
        {
            userInput.FilePath = WorkbookPath.Text;
            userInput.TemplateNames = TemplateRange.Text;
            userInput.SourceNames = SourceRange.Text;
        }

        private void CheckWorksheets()
        {
            ExcelDialogs.CheckSheets(userInput.FilePath);

            // Only the status of the buttons is updated here, in contrast to other scripts.
            // I'm not sure what kind of autofill behaviour would be suitable here.
            Dispatcher.Invoke(UpdateButtons);

            void UpdateButtons()
            {
                SelectTemplates.IsEnabled = ExcelDialogs.SheetRangeDialogAvailable;
                SelectSources.IsEnabled = SelectTemplates.IsEnabled;
            }
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
