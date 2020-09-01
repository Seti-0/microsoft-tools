using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

using Red.Core.Logs;
using System.Diagnostics;
using Red.Core.IO;

namespace Red.Core
{
    using ExcelApp = Microsoft.Office.Interop.Excel.Application;
    using WordApp = Microsoft.Office.Interop.Word.Application;

    public class OfficeApps
    {
        public static void RunOfficeWithGuard(Action<OfficeApps> action)
        {
            RunAndCollect(true, true, action);
        }

        public static void RunExcelWithGuard(Action<OfficeApps> action)
        {
            RunAndCollect(true, false, action);
        }

        public static void RunWordWithGuard(Action<OfficeApps> action)
        {
            RunAndCollect(false, true, action);
        }

        private static void RunAndCollect(bool excel, bool word, Action<OfficeApps>  action)
        {
            // In debug mode, a local variable is considered used for the duration of it's method, 
            // for the sake of being able to see it, and so cannot be collected. For GC to
            // have any effect in that case, the OfficeApps variable has to be kept in a separate method.
            RunWithGuard(excel, word, action);

            // As far as I can tell, some (all?) of the interop objects need to have been collected
            // (and perhaps their finalizers run?) before the external process will terminate naturally. 
            // This doesn't happen if this program exits before another collection is needed, so this has to be done
            // manually.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void RunWithGuard(bool excel, bool word, Action<OfficeApps> action)
        {
            OfficeApps apps = null;

            try
            {
                apps = new OfficeApps();

                if (excel)
                    apps.StartExcel();

                if (word)
                    apps.StartWord();

                action(apps);
            }
            catch (Exception e)
            {
                Log.Core.Error("Unexpected error occurred, aborting action", e);
            }
            finally
            {
                // I really should come back to exceptions here, it's silly and messy at the moment.
                
                // Currently, this block can throw an exception, causing the original exception to be lost.
                // To make matters worse, logging this exception could also cause and exception the way things are.

                // At the end of the day, an exception here should probably lead to a fallback of terminating
                // the stray process?

                if (apps != null)
                    apps.Quit();
            }
        }

        private ExcelApp excel = null;
        private WordApp word = null;

        public ExcelApp Excel
        {
            get
            {
                if (excel == null)
                    throw new InvalidOperationException("Attempted to use Excel with starting it");

                return excel;
            }

            private set => excel = value;
        }

        public WordApp Word
        {
            get
            {
                if (word == null)
                    throw new InvalidOperationException("Attempted to use Word without starting it");

                return word;
            }

            private set => word = value;
        }

        public ExcelApp StartExcel()
        {
            if (excel == null)
            {
                Log.Core.Debug("Starting Excel");
                excel = new ExcelApp();
            }
            
            return excel;
        }

        public WordApp StartWord()
        {
            if (word == null)
            {
                Log.Core.Debug("Starting Word");
                word = new WordApp();
            }
            
            return word;
        }

        /// <summary>
        /// This closes the excel and word apps if they are open, 
        /// but this is not enough to cleanup the processes. For that, the GC "collect" and
        /// "wait for finalizers" must be called. Furthermore, in debug mode, they must be called
        /// after any method using this object has returned. This means that this object cannot be used
        /// safely in the main method. It also means this object cannot be kept as a static field, since it
        /// would then never be released.
        /// </summary>
        public void Quit()
        {
            if (word != null)
            {
                Log.Core.Debug("Closing Word");

                foreach (Document document in word.Documents)
                    if (!document.Saved)
                        Log.Core.Warning($"Discarding changes to {document.FullName}");

                word.Quit(SaveChanges: false);
                word = null;
            }

            if (excel != null)
            {
                excel.DisplayAlerts = false;

                foreach (Workbook workbook in excel.Workbooks)
                {
                    if (!workbook.Saved)
                        Log.Core.Warning($"Discarding changes to {workbook.FullName}");

                    workbook.Close(SaveChanges: false);
                }

                Log.Core.Debug("Closing Excel");
                excel.Quit();

                // I don't like this, but until I can figure out something better it will stay.
                KillExcel();

                excel = null;
            }

            FileHelper.DeleteTemporaryFiles();
        }

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private void KillExcel()
        {
            /* 
             * This is a last resort. Aside from just not being nice, there
             * are obscure situations in which nasty side effects can occur, apparently. (Something
             * about shared excel processess, I didn't understand it)
             */

            GetWindowThreadProcessId(excel.Hwnd, out int id);
            var excelProc = Process.GetProcessById(id);
            excelProc.Kill();
        }
    }
}
