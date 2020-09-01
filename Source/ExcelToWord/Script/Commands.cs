using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Red.Core.Office;

namespace ExcelToWord
{
    using WordRange = Microsoft.Office.Interop.Word.Range;
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;

    public class CommandContext
    {
        public string Name;

        public Dictionary<string, Workbook> BooksByAlias = new Dictionary<string, Workbook>();
        public Dictionary<string, Workbook> BooksByName = new Dictionary<string, Workbook>();

        public bool CheckID(string bookID)
        {
            return bookID == null || BooksByAlias.ContainsKey(bookID) || BooksByName.ContainsKey(bookID);
        }

        public Worksheet GetSheet(string bookID)
        {
            Workbook book;

            if (bookID == null)
                book = BooksByName.Values.First();

            else if (!BooksByAlias.TryGetValue(bookID, out book))
            {
                if (!BooksByName.TryGetValue(bookID, out book))
                {
                    Script.Log.Warning(
                        $"Could not find excel source with alias or name \"{bookID}\"");

                    return null;
                }
            }

            if (!ExcelHelper.TrySelectWorksheet(book, out Worksheet result, Name,
                compareWords: true, verbrose: false))
            {
                Script.Log.Warning($"Could not find sheet \"{Name}\" in workbook {book.Name}");
            }

            return result;
        }
    }

    public interface ICommand
    {
        bool Check(CommandContext context);

        void Apply(CommandContext context);
    }

    public class SheetNameReference : ICommand
    {
        private readonly WordRange target;

        public SheetNameReference(WordRange target)
        {
            this.target = target;
        }

        public bool Check(CommandContext context) => true;

        public void Apply(CommandContext context)
        {
            target.Text = context.Name.ToUpper();
        }
    }

    /*
    public class ClearCommand : ICommand
    {
        private readonly WordRange target;

        public ClearCommand(WordRange target)
        {
            this.target = target;
        }

        public void Apply(CommandContext context)
        {
            target.Text = "";
        }
    }
    */

    public class DirectReference : ICommand
    {
        private readonly WordRange target;
        private readonly string cellReference;
        private readonly string workbookID;

        public DirectReference(string workbookID, WordRange target, string cellReference)
        {
            this.target = target;
            this.cellReference = cellReference;
            this.workbookID = workbookID;
        }

        public bool Check(CommandContext context)
        {
            bool passed = context.CheckID(workbookID);

            if (!passed)
            {
                Script.Log.Warning($"Could not find workbook with name or alias {workbookID} among sources");
                Script.Log.Debug($"Skipping reference to {workbookID}: {cellReference}");
                return false;
            }

            return true;
        }

        public void Apply(CommandContext context)
        {
            Worksheet sheet = context.GetSheet(workbookID);

            if (sheet == null)
                return;

            ExcelRange source = null;

            try
            {
                source = sheet.Cells.Range[cellReference];
            }
            catch (Exception e)
            {
                Script.Log.Warning($"Failed to retrieve Cell Reference {cellReference} from Worksheet {sheet.Name}", e);
                Script.Log.Debug("Is it a valid reference?");
            }

            try
            {
                if (source != null)
                    target.Text = RangeToText(source);
            }
            catch (Exception e)
            {
                Script.Log.Warning($"Failed to apply Cell Reference {cellReference} from Worksheet {sheet.Name}", e);
            }
        }

        private static string RangeToText(ExcelRange range)
        {
            IList<string> items = new List<string>();

            RangeEnumerator enumerator = new RangeEnumerator(range);
            while(enumerator.MoveNext())
            {
                ExcelRange cell = enumerator.Current;
                if (cell != null && !string.IsNullOrWhiteSpace(cell.Text?.ToString()))
                {
                    string text = cell.Text?.ToString();
                    items.Add(text);
                }
            }

            if (items.Count > 0)
                return string.Join("\n", items);
            else
                return "";
        }
    }
}
