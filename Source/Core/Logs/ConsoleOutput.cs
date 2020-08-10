using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace Red.Core.Logs
{

    public static class ConsoleOutput
    {
        private static readonly string splitPattern = @"\.|,|\/|\\|\[|\]|\{|\}|\(|\)|\s+";

        private struct ColorScheme
        {
            public ConsoleColor Highlight;
            public ConsoleColor Normal;

            public ConsoleColor DarkHighlight;
            public ConsoleColor Dark;
        }

        private static readonly Dictionary<Log.Type, ColorScheme> colorSchemes = new Dictionary<Log.Type, ColorScheme>()
        {
            {
                Log.Type.Debug, new ColorScheme
                {
                    Highlight = ConsoleColor.White,
                    Normal = ConsoleColor.Gray,
                    DarkHighlight = ConsoleColor.DarkGray,
                    Dark = ConsoleColor.DarkGray,
                }
            },
            {
                Log.Type.Fine, new ColorScheme
                {
                    Highlight = ConsoleColor.Gray,
                    Normal = ConsoleColor.DarkGray,
                    DarkHighlight = ConsoleColor.DarkGray,
                    Dark = ConsoleColor.DarkGray,
                }
            },
            {
                Log.Type.Info, new ColorScheme
                {
                    Highlight = ConsoleColor.Cyan,
                    Normal = ConsoleColor.Blue,
                    DarkHighlight = ConsoleColor.DarkCyan,
                    Dark = ConsoleColor.DarkBlue,
                }
            },
            {
                Log.Type.Warning, new ColorScheme
                {
                    Highlight = ConsoleColor.Red,
                    Normal = ConsoleColor.Yellow,
                    DarkHighlight = ConsoleColor.DarkRed,
                    Dark = ConsoleColor.DarkYellow,
                }
            },
            {
                Log.Type.Error, new ColorScheme
                {
                    Highlight = ConsoleColor.Yellow,
                    Normal = ConsoleColor.Red,
                    DarkHighlight = ConsoleColor.DarkYellow,
                    Dark = ConsoleColor.DarkRed,
                }
            },
            {
                Log.Type.Success, new ColorScheme
                {
                    Highlight = ConsoleColor.Yellow,
                    Normal = ConsoleColor.Green,
                    DarkHighlight = ConsoleColor.DarkGreen,
                    Dark = ConsoleColor.DarkYellow,
                }
            },
            {
                Log.Type.Special, new ColorScheme
                {
                    Highlight = ConsoleColor.Cyan,
                    Normal = ConsoleColor.Magenta,
                    DarkHighlight = ConsoleColor.DarkCyan,
                    Dark = ConsoleColor.DarkMagenta,
                }
            },
            {
                Log.Type.Aside, new ColorScheme
                {
                    Highlight = ConsoleColor.Cyan,
                    Normal = ConsoleColor.Magenta,
                    DarkHighlight = ConsoleColor.DarkGray,
                    Dark = ConsoleColor.DarkMagenta,
                }
            },
        };

        private class Memory
        {
            public string Variable = "", Fixed = "", Date = "";
        }

        private static Dictionary<string, Memory> memories = new Dictionary<string, Memory>();
        private static IConsole console;
        private static Memory currentMemory;

        private static string nextFixed, nextVariable, nextDate;

        private static void ApplyIndent(int indent)
        {
            if (indent < 0)
                indent = 0;

            console.Write(new string('\t', indent));
        }

        public static void Write(IConsole console, Log.Entry entry, string category)
        {
            ConsoleOutput.console = console;

            /*if (memories.ContainsKey(console.ID))
                currentMemory = memories[console.ID];

            else
            {
                currentMemory = new Memory();
                //memories.Add(console.ID, currentMemory);
            }*/

            currentMemory = new Memory();

            nextFixed = "";
            nextVariable = "";
            nextDate = "";

            ColorScheme colorScheme;

            if (!colorSchemes.TryGetValue(entry.Type, out colorScheme))
                colorScheme = colorSchemes[Log.Type.Debug];

            ConsoleOutput.console.ForegroundColor = ConsoleColor.DarkGray;

            //string date = entry.Time.ToShortDateString();
            //if (date != currentMemory.Date)
            //    ConsoleOutput.console.WriteLine($"[{date}]");

            //nextDate = date;

            ConsoleOutput.console.Write($"[{entry.Time.ToShortTimeString()}]");
            ConsoleOutput.console.Write($"[{category}] ");
            ApplyIndent(entry.Indent);

            string formatItemPattern = @"\{((\d+):)?[^\{]*\}";

            var matches = Regex.Matches(entry.Message, formatItemPattern);

            int currentIndex = 0;
            foreach (Match match in matches)
                if (match.Success)
                {
                    string fixedSection = entry.Message.Substring(currentIndex, match.Index - currentIndex);
                    string formatSection = entry.Message.Substring(match.Index, match.Length);
                    currentIndex = match.Index + match.Length;

                    WriteFixed(fixedSection, colorScheme);
                    WriteVariable(match, entry.Parameters, colorScheme);
                }

            string finalSection = entry.Message.Substring(currentIndex);
            WriteFixed(finalSection, colorScheme);

            ConsoleOutput.console.WriteLine();

            WriteException(entry.Exception, entry.Indent, entry.Type != Log.Type.Fine);

            currentMemory.Fixed = nextFixed;
            currentMemory.Variable = nextVariable;
            currentMemory.Date = nextDate;
        }

        private static bool Split(string input, string pattern, out string upto, out string after)
        {
            var match = Regex.Match(input, pattern);

            if (match.Success)
            {
                upto = input.Substring(0, match.Index + 1);
                after = input.Substring(match.Index + 1);
                return true;
            }
            else
            {
                upto = null;
                after = null;
                return false;
            }
        }

        private static void WriteVariable(Match formatSection, object[] parameters, ColorScheme colorScheme)
        {
            if (formatSection.Groups.Count > 2)
            {
                if (formatSection.Groups.Count > 2 && int.TryParse(formatSection.Groups[2].Value, out int parameterIndex)
                        && parameters != null && parameters.Length > parameterIndex)
                {
                    WriteVariable(formatSection.Value, parameters, colorScheme);
                }
                else if (formatSection.Groups[2].Value == "")
                {
                    WriteVariable(formatSection.Value, parameters, colorScheme);
                }
            }
            else
            {
                WriteFixed(formatSection.Value, colorSchemes[Log.Type.Error]);
            }

        }

        private static void WriteFixed(string fixedSection, ColorScheme colorScheme)
        {
            while (Split(fixedSection, splitPattern, out string upto, out string after))
            {
                Write(upto, colorScheme.Normal, colorScheme.Dark, currentMemory.Fixed, ref nextFixed);

                fixedSection = after;
            }

            Write(fixedSection, colorScheme.Normal, colorScheme.Dark, currentMemory.Fixed, ref nextFixed);
        }

        private static void WriteVariable(string formatSection, object[] parameters, ColorScheme colorScheme)
        {
            var variableSection = string.Format(formatSection, parameters);

            while (Split(variableSection, splitPattern, out string upto, out string after))
            {
                Write(upto, colorScheme.Highlight, colorScheme.DarkHighlight, currentMemory.Variable, ref nextVariable);

                variableSection = after;
            }

            Write(variableSection, colorScheme.Highlight, colorScheme.DarkHighlight, currentMemory.Variable, ref nextVariable);
        }

        private static void Write(string text, ConsoleColor light, ConsoleColor dark, string referenceText, ref string nextReference)
        {
            if (referenceText.Contains(text))
                console.ForegroundColor = dark;
            else
                console.ForegroundColor = light;

            console.Write(text);
            nextReference += text;
        }

        private static void WriteException(Exception e, int indent, bool color = true)
        {
            if (e == null)
                return;

            var primary = color ? ConsoleColor.Red : ConsoleColor.DarkGray;
            var secondary = color ? ConsoleColor.Yellow : ConsoleColor.DarkGray;

            string name = e.GetType().Name;
            string message = e.Message;

            console.ForegroundColor = primary;
            ApplyIndent(indent);
            console.WriteLine($"[{name}] {message}");

            var trace = new StackTrace(e, true).GetFrames();

            if (color)
            {
                primary = ConsoleColor.DarkGray;
                secondary = ConsoleColor.DarkRed;
            }

            if (trace == null)
            {
                console.ForegroundColor = ConsoleColor.DarkGray;
                ApplyIndent(indent);
                console.WriteLine($"(Stacktrace is null)");
            }
            else
            {
                bool missingFilenames = false;
                bool missingLinenumbers = false;

                foreach (var frame in trace)
                {
                    ApplyIndent(indent + 1);

                    SplitMethodSignature(frame.GetMethod().ToString(), out string methodName, out string arguments);

                    string fileName = frame.GetFileName();
                    bool missingFilename = string.IsNullOrWhiteSpace(fileName);
                    missingFilenames |= missingFilename;

                    int lineNumber = frame.GetFileLineNumber();
                    bool missingLineNumber = lineNumber == 0;
                    missingLinenumbers |= missingLineNumber;

                    console.ForegroundColor = primary;
                    console.Write($"at ");

                    if (!missingLineNumber)
                    {
                        console.Write("line ");
                        console.ForegroundColor = secondary;
                        console.Write(lineNumber.ToString());
                        console.ForegroundColor = primary;
                        console.Write(" of ");
                    }

                    console.ForegroundColor = secondary;
                    console.Write(methodName);

                    console.ForegroundColor = primary;
                    console.Write($"{arguments}");

                    if (!missingFilename)
                    {
                        console.Write($"in {frame.GetFileName()}");
                    }

                    console.WriteLine();
                }

                if (missingLinenumbers || missingFilenames)
                {
                    string infoMissing = null;

                    if (missingFilenames && missingLinenumbers)
                        infoMissing = "file names/line numbers";
                    else if (missingFilenames)
                        infoMissing = "file names";
                    else
                        infoMissing = "line numbers";

                    console.ForegroundColor = primary;
                    console.WriteLine($"(Some {infoMissing} missing, perhaps because Release config. is involved?)");
                }
            }

            if (e.InnerException != null)
            {
                console.ForegroundColor = primary;
                ApplyIndent(indent);
                console.WriteLine("Caused by Inner Exception:");
                WriteException(e.InnerException, indent, color);
            }
        }

        private static void SplitMethodSignature(string input, out string name, out string arguments)
        {
            int argumentIndex = input.IndexOf('(');
            int genericIndex = input.IndexOf('<');

            int index = input.Length - 1;

            if (argumentIndex > -1)
                index = argumentIndex;

            if (genericIndex > -1 && genericIndex < argumentIndex)
                index = genericIndex;

            name = input.Substring(0, index);
            arguments = input.Substring(index);
        }
    }
}
