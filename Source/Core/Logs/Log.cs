using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;

namespace Red.Core.Logs
{
    public class Log : IEnumerable<Log.Entry>
    {
        private static object _lockObj = new object();

        public static List<IConsole> Outputs { get; set; } = new List<IConsole>()
        {
            new StandardConsole(),
        };

        public static void EndOutputs()
        {
            foreach (var console in Outputs)
                console.End();
        }

        public static Log Core { get; } = new Log("Core");

        public enum Type
        {
            Debug, Fine, Info, Warning, Error,
            Success, Aside, Special
        }

        public class Entry
        {
            public string Message;
            public object[] Parameters;

            public Type Type;
            public Exception Exception;
            public DateTime Time = DateTime.Now;
            public int Indent;
        }

        public string Title { get; }

        private readonly IList<Entry> entries = new List<Entry>();
        private int indent;

        public Log(string title = "Core")
        {
            Title = title;
        }

        public void ResetIndent()
        {
            indent = 0;
        }

        public void PushIndent()
        {
            indent++;
        }

        public void PopIndent()
        {
            indent--;

            if (indent < 0)
                indent = 0;
        }

        public void Append(Entry entry)
        {
            lock(_lockObj)
            {
                entries.Add(entry);

                foreach (var output in Outputs)
                    ConsoleOutput.Write(output, entry, Title);
            }
        }

        public IEnumerator<Entry> GetEnumerator()
        {
            return entries.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return entries.GetEnumerator();
        }

        #region Convenience
        public void Write(Log.Type type, string message, params object[] parameters)
        {
            Append(new Log.Entry { Type = type, Message = message, Parameters = parameters, Indent = indent });
        }

        public void Write(Log.Type type, string message, Exception e, params object[] parameters)
        {
            Append(new Log.Entry { Type = type, Message = message, Parameters = parameters, Indent = indent, Exception = e });
        }

        public void Debug(string message) => Write(Log.Type.Debug, message);

        public void Fine(string message, params object[] parameters) => Write(Log.Type.Fine, message, parameters);
        public void Debug(string message, params object[] parameters) => Write(Log.Type.Debug, message, parameters);
        public void Info(string message, params object[] parameters) => Write(Log.Type.Info, message, parameters);
        public void Warning(string message, params object[] parameters) => Write(Log.Type.Warning, message, parameters);
        public void Error(string message, params object[] parameters) => Write(Log.Type.Error, message, parameters);
        public void Success(string message, params object[] parameters) => Write(Log.Type.Success, message, parameters);
        public void Aside(string message, params object[] parameters) => Write(Log.Type.Aside, message, parameters);
        public void Special(string message, params object[] parameters) => Write(Log.Type.Special, message, parameters);

        public void Debug(string message, Exception e, params object[] parameters) => Write(Log.Type.Debug, message, e, parameters);
        public void Warning(string message, Exception e, params object[] parameters) => Write(Log.Type.Warning, message, e, parameters);
        public void Error(string message, Exception e, params object[] parameters) => Write(Log.Type.Error, message, e, parameters);
        public void Special(string message, Exception e, params object[] parameters) => Write(Log.Type.Special, message, e, parameters);
        #endregion
    }
}