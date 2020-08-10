using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Red.Core.Logs
{
    // I know, this is backwards. The file output certainly should not need to implement IConsole, and ignore
    // color. But this happens to be the easist way to do this right now.
    public class FileConsole : IConsole
    {
        public static bool Active { get; private set; } = false;

        public static void Activate()
        {
            if (Active)
                return;

            Log.Outputs.Add(new FileConsole("File"));
            Active = true;
        }

        public string ID { get; }

        public ConsoleColor ForegroundColor { get; set; }

        public string FilePath { get; }

        public FileConsole(string id = "File", string path = "log.txt")
        {
            ID = id;
            FilePath = path;

            WriteLine($"==========[ Log started at {DateTime.Now} ]==========");
        }

        public void Write(string text)
        {
            try
            {
                File.AppendAllText(FilePath, text);
            }
            catch(Exception e)
            {
                // For now, this is the best that can be done. Ideally, a failure here
                // should prevent other outputs from working, though
                Console.WriteLine("Failed to write to logfile");
                Console.WriteLine(e.ToString());
            }
        }

        public void WriteLine(string text = "")
        {
            Write(text + "\n");
        }

        public void End()
        {
            WriteLine($"==========[ Log ended at {DateTime.Now} ]==========");
        }
    }
}
