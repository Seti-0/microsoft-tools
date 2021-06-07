using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Controls;
using Red.Core;

namespace WpfToolset
{
    /// <summary>
    /// Interaction logic for SelectTemplateDialog.xaml
    /// </summary>
    public partial class SelectWorksheetDialog : Window
    {
        public class Item : INotifyPropertyChanged
        {
            private string content;
            private bool primary;
            
            public object UserData { get; set; }

            public string Content
            {
                get => content;

                set
                {
                    if (content != value)
                    {
                        content = value;
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Content)));
                    }
                }
            }

            public bool Primary
            {
                get => primary;

                set
                {
                    if (primary != value)
                    {
                        primary = value;
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Primary)));
                    }
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
        }

        private static SelectWorksheetDialog lastDialog;

        public static SelectWorksheetDialog Current
        {
            get
            {
                if (lastDialog != null && lastDialog.IsActive)
                    return lastDialog;

                else return null;
            }
        }

        public string WorkbookPath { get; private set; }

        public string Result { get; private set; }

        public List<string> Results { get; private set; }

        public IEnumerable<string> ItemContent
        {
            get => Items.Select(x => x.Content);
            set => Items = value.Select(x => new Item { Content = x, Primary = true});
        }

        public IEnumerable<Item> Items
        {
            get => List.ItemsSource.Cast<Item>();
            set => List.ItemsSource = value;
        }

        public SelectWorksheetDialog(string workbookPath, SelectionMode selectionMode)
        {
            InitializeComponent();
            WorkbookPath = workbookPath;
            List.SelectionMode = selectionMode;

            lastDialog = this;

            if (workbookPath == null)
                Refresh.IsEnabled = false;

            SelectAll(primaryOnly: true);
        }

        public void SelectAll(bool primaryOnly = false)
        {
            switch (List.SelectionMode)
            {
                case SelectionMode.Single:
                    
                    if (!primaryOnly)
                    {
                        List.SelectedIndex = 0;
                        break;
                    }

                    List.SelectedItems.Clear();
                    foreach (object obj in List.Items)
                        if (obj is Item item && item.Primary)
                        {
                            List.SelectedItems.Add(obj);
                            break;
                        }
                    
                    break;

                case SelectionMode.Extended:
                case SelectionMode.Multiple:

                    List.SelectedItems.Clear();

                    foreach (object obj in List.Items)
                        if ((!primaryOnly) || obj is Item item && item.Primary)
                            List.SelectedItems.Add(obj);

                    break;
            }

            if (List.SelectedItems.Count > 0)
                List.ScrollIntoView(List.SelectedItems[0]);
        }

        public void SetSelectedRange(int a, int b)
        {
            if (a < 0)
                throw new ArgumentOutOfRangeException(nameof(a));

            if (a >= List.Items.Count)
                throw new ArgumentOutOfRangeException(nameof(a));

            if (b < a)
                throw new ArgumentOutOfRangeException(nameof(b));

            if (b > List.Items.Count)
                throw new ArgumentOutOfRangeException(nameof(b));

            for (int i = a; i < b; i++)
            {
                List.SelectedItems.Add(List.Items[i]);
            }
        }

        private void Okay_Click(object sender, RoutedEventArgs e)
        {
            Result = GetResult(List.SelectedItem as Item);
            Results = List.SelectedItems.Cast<Item>().Select(GetResult).ToList();
            DialogResult = Result != null;
        }

        private string GetResult(Item item)
        {
            if (item == null)
                return "<null>";

            if (item.UserData != null)
                return item.UserData.ToString();

            else return item.Content?.ToString() ?? "<null>";
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            RunWithCancel("Refresh Worksheets", () => ExcelDialogs.CheckSheets(WorkbookPath), "Dialog refresh cancelled");
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Flow.Interrupt(Flow.Reason.Cancel);
        }

        private async void RunWithCancel(string name, Action action, string cancelMessage)
        {
            Refresh.Content = "Cancel";
            Refresh.Click -= Refresh_Click;
            Refresh.Click += Cancel_Click;

            await System.Threading.Tasks.Task.Run(() => Flow.RunWithCancel(name, action, cancelMessage));

            if (Flow.Interrupted && Flow.InterruptReason == Flow.Reason.Quit)
                Application.Current.Shutdown();

            Refresh.Content = "Refresh";
            Refresh.Click -= Cancel_Click;
            Refresh.Click += Refresh_Click;
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void CommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SelectAll();
        }
    }
}
