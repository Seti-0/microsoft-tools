using System;
using System.Collections.Generic;
using System.Text;

namespace Red.Core.Logs
{
    public interface IConsole
    {
        string ID { get; }

        ConsoleColor ForegroundColor { get; set; }

        void Write(string text);

        void WriteLine(string text = "");

        void End();
    }
}
