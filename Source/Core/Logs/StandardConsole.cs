using System;
using System.Collections.Generic;
using System.Text;

namespace Red.Core.Logs
{
    public class StandardConsole : IConsole
    {
        public string ID { get; }

        public StandardConsole(string name = "Standard")
        {
            ID = name;
        }

        public ConsoleColor ForegroundColor
        {
            get => Console.ForegroundColor;

            set => Console.BackgroundColor = value;
        }

        public void Write(string text)
        {
            Console.Write(text);
        }

        public void WriteLine(string text = "")
        {
            Console.WriteLine(text);
        }

        public void Begin(){}

        public void End(){}
    }
}
