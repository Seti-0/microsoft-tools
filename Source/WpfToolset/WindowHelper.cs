using System;
using System.Collections.Generic;
using System.Text;

using Red.Core;
using Red.Core.Logs;

using System.Windows;
using System.Windows.Controls;

namespace WpfToolset
{
    public static class WindowHelper
    {
        public static Button Run { get; set; }

        public static Button Cancel { get; set; }

        public static void Setup(Window window, RichTextBox logBox, Button run, Button cancel)
        {
            var console = new TextBoxConsole("TextBoxConsole", logBox);
            Log.Outputs.Add(console);

            Run = run;
            Cancel = cancel;

            Cancel.IsEnabled = false;

            window.Closed += Window_Closed;
        }

        private static void Window_Closed(object sender, EventArgs e)
        {
            if (Flow.Idle)
                AppHelper.Shutdown("App Idle, Shutting Down");

            else if (!AppHelper.EarlyShutdown)
                Flow.Interrupt(Flow.Reason.Quit);
        }

        public static async void RunWithCancel(string name, Action action, string cancelMessage)
        {
            Run.IsEnabled = false;
            Cancel.IsEnabled = true;

            await System.Threading.Tasks.Task.Run(() => Flow.RunWithCancel(name, action, cancelMessage));

            if (Flow.Interrupted && Flow.InterruptReason == Flow.Reason.Quit)
                AppHelper.Shutdown("Operation interruption complete. Reason: User Exit. Shutting down");

            Run.IsEnabled = true;
            Cancel.IsEnabled = false;
        }
    }
}
