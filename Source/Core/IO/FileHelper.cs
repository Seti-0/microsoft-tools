using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Red.Core.Logs;

namespace Red.Core.IO
{
    public class FileHelper
    {
        private static List<string> _temporaryPaths = new List<string>();

        public static bool TryOpenWorkbook(OfficeApps apps, string filePath, 
            bool readOnly, out Workbook result)
        {
            var app = apps.Excel;
            result = null;
            bool found = false;

            if (Flow.Interrupted)
                return false;

            try
            {
                result = app.Workbooks.Open(filePath, ReadOnly: readOnly);
                Log.Core.Debug($"Opening {result.Name}");
                found = true;
            }
            catch (Exception e)
            {
                Log.Core.Error($"Failed to read workbook at {filePath}", e);
            }

            return found;
        }

        public static bool TryOpenDocument(OfficeApps apps, string filePath,
            bool readOnly, out Document result)
        {
            var app = apps.Word;
            result = null;
            bool found = false;

            if (Flow.Interrupted)
                return false;

            try
            {
                result = app.Documents.Open(filePath, ReadOnly: readOnly);
                Log.Core.Debug($"Opening {result.Name}");
                found = true;
            }
            catch (Exception e)
            {
                Log.Core.Error($"Failed to read document at {filePath}", e);
            }

            return found;
        }

        public static void SaveTemporarily(Document document)
        {
            string path = PathHelper.GetUniqueFileName("temporary file.docx");
            path = Path.GetFullPath(path);
            _temporaryPaths.Add(path);

            document.SaveAs2(path);
            Log.Core.Debug($"Creating temporary file: {path}");
        }

        public static void DeleteTemporaryFiles()
        {
            if (_temporaryPaths.Count == 0)
                return;

            Log.Core.Debug("Deleting temporary files");

            foreach (string path in _temporaryPaths)
            {
                try
                {
                    File.Delete(path);
                }
                catch (Exception e)
                {
                    Log.Core.Warning("Failed to delete temporary file");
                    Log.Core.Debug("This means it may have to be deleted manually", e);
                }
            }

            _temporaryPaths.Clear();
        }
    }
}
