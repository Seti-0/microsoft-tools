using System;
using System.Collections.Generic;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Red.Core.Logs;

namespace Red.Core.IO
{
    public class FileHelper
    {
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

        public static Dictionary<string, Workbook> OpenWorkbooks(OfficeApps apps, 
            IList<string> filePaths, bool readOnly)
        {
            var results = new Dictionary<string, Workbook>();
            foreach (var path in filePaths)
            {
                if (Flow.Interrupted)
                    break;

                if (TryOpenWorkbook(apps, path, readOnly, out Workbook book))
                    results.Add(book.Name, book);
            }
            return results;
        }

        public static Dictionary<string, Document> OpenDocuments(OfficeApps apps,
            IList<string> filePaths, bool readOnly)
        {
            var results = new Dictionary<string, Document>();
            foreach (var path in filePaths)
            {
                if (Flow.Interrupted)
                    break;

                if (TryOpenDocument(apps, path, readOnly, out Document doc))
                    results.Add(doc.Name, doc);
            }
            return results;
        }
    }
}
