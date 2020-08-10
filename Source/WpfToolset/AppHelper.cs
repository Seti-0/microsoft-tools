using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Threading.Tasks;
using Red.Core;
using Red.Core.Logs;

namespace WpfToolset
{
    public static class AppHelper
    {
        public static bool DebugFlag { get; private set; } = false;

        public static bool EarlyShutdown { get; private set; }

        public static void AttachErrorHandling(Application app, StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

#if DEBUG
            DebugFlag = true;
#endif
            if (e.Args.Length > 0)
            {
                var arg = e.Args[0];
                if (arg == "-d" || arg == "--debug")
                {
                    DebugFlag = true;
                }
            }

            if (DebugFlag)
                FileConsole.Activate();

            app.DispatcherUnhandledException += App_DispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        public static void Shutdown(string bye)
        {
            Log.Core.Debug(bye);
            Log.EndOutputs();
            Application.Current.Shutdown();
        }

        private static void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            HandleException(e.Exception);
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            // Not sure if this is right
            HandleException(e.ExceptionObject as Exception);
        }

        private static void App_DispatcherUnhandledException(object sender, System.Windows.Threading
            .DispatcherUnhandledExceptionEventArgs e)
        {
            HandleException(e.Exception);
            e.Handled = true;
        }

        private static void HandleException(Exception e)
        {
            Logs.Wpf.Error("Unknown error occurred", e);

            if (!Flow.Initialized)
            {
                EarlyShutdown = true;
                Shutdown("Error occured during program initialization. Shutting down");
            }
        }
    }
}
