using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Navigation;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Red.Core;
using Red.Core.IO;
using Red.Core.Logs;
using Red.Core.Office;
using WpfToolset;

namespace ExcelToWord
{
    using ExcelRange = Microsoft.Office.Interop.Excel.Range;
    using WordRange = Microsoft.Office.Interop.Word.Range;

    public class Script
    {
        public static Log Log { get; } = new Log("Script");

        public static void Execute(OfficeApps apps, Input input)
        {
            if (Flow.Interrupted)
                return;

            Log.Info("Executing Script");

            Log.Debug("Template: " + input.Template.Name);

            string outputName = Path.GetFileNameWithoutExtension(input.Template.Name);
            string outputExtension;
            WdSaveFormat outputFormat;

            switch (input.OutputFormat)
            {
                case OutputFormat.Word:
                    outputExtension = "docx";
                    outputFormat = WdSaveFormat.wdFormatDocument;
                    break;

                case OutputFormat.PDF:
                    outputExtension = "pdf";
                    outputFormat = WdSaveFormat.wdFormatPDF;
                    break;

                default:
                    
                    string name = Enum.GetName(typeof(OutputFormat), input.OutputFormat);
                    Log.Warning($"Unrecognized output format {name}");
                    Log.Debug("Defaulting to PDF");

                    outputExtension = "pdf";
                    outputFormat = WdSaveFormat.wdFormatPDF;
                    break;
            }



            if (input.Sources.Count == 0)
            {
                Log.Warning("No sources found. Aborting.");
                return;
            }

            Log.Debug("Sources:");
            Log.PushIndent();
            foreach (Input.Source source in input.Sources)
            {
                Log.Debug($"{source.Alias} - {source.Workbook.Name} ({source.Workbook.Path})");

                if (Flow.Interrupted)
                    return;
            }
            Log.PopIndent();

            if (input.SheetNames.Count == 0)
            {
                Log.Warning("No sheet names found. Aborting");
                return;
            }

            Log.Debug("Sheet Names:");
            Log.PushIndent();
            foreach (string name in input.SheetNames)
            {
                Log.Debug(name);

                if (Flow.Interrupted)
                    return;
            }
            Log.PopIndent();

            if (Flow.Interrupted)
                return;

            FileHelper.SaveTemporarily(input.Template);

            Log.Info($"Reading template");

            List<ICommand> commands = CollectCommands(input.Template);

            if (!commands.Any())
            {
                Log.Warning($"No commands found. Aborting.");
                return;
            }

            else Log.Debug($"{commands.Count()} found");

            CommandContext context = new CommandContext();
            foreach (Input.Source source in input.Sources)
            {
                context.BooksByAlias.Add(source.Alias, source.Workbook);
                context.BooksByName.Add(source.Workbook.Name, source.Workbook);
            }

            Log.Debug("Checking commands");
            List<ICommand> checkedCommands = commands
                .Where(x => x.Check(context))
                .ToList();

            if (!checkedCommands.Any())
            {
                Log.Warning($"No commands passed check. Aborting.");
                return;
            }

            Log.Debug($"Applying template to:");
            Log.PushIndent();

            if (Flow.Interrupted)
                return;

            foreach (string sheetName in input.SheetNames)
            {
                context.Name = sheetName;
                Log.Debug(sheetName);

                foreach (var command in checkedCommands)
                {
                    command.Apply(context);

                    if (Flow.Interrupted)
                        break;
                }

                if (Flow.Interrupted)
                    break;

                string parent = Path.GetFullPath("Output");
                Directory.CreateDirectory(parent);

                string fileName = PathHelper.CreateValidFilename($"{outputName} - {sheetName}.{outputExtension}");
                string filePath = Path.Combine(parent, PathHelper.GetUniqueFileName(fileName));

                input.Template.SaveAs2(filePath, outputFormat);
            }

            Log.PopIndent();

            if (Flow.Interrupted)
                return;

            // This is saving to the temporary file, not overwriting anything.
            input.Template.Save();

            Log.Success("Script complete");
        }

        private static List<ICommand> CollectCommands(Document template)
        {
            template.Activate();

            WordRange searchRange = template.Range();
            searchRange.Find.Text = "#<*>#";
            searchRange.Find.MatchWildcards = true;
            searchRange.Find.MatchWholeWord = true;

            var results = new List<ICommand>();

            while (searchRange.Find.Execute())
            {
                var command = CreateCommand(searchRange.Duplicate);

                if (command != null)
                    results.Add(command);
            }

            return results;
        }

        private static ICommand CreateCommand(WordRange commandRange)
        {
            string sheetnamePattern = "#[Ss][Hh][Ee][Ee][Tt][Nn][Aa][Mm][Ee]#";
            string directReferencePattern = @"#(([^-]+)-)?([^#]+)#";

            Match match;

            match = Regex.Match(commandRange.Text, sheetnamePattern);
            if (match.Success)
                return new SheetNameReference(commandRange);

            match = Regex.Match(commandRange.Text, directReferencePattern);
            if (match.Success)
            {
                string bookID = match.Groups[2].Value;
                string cellReference = match.Groups[3].Value;

                if (string.IsNullOrWhiteSpace(bookID))
                    bookID = null;

                return new DirectReference(bookID, commandRange, cellReference);
            }

            return null;
        }
    }
}
