using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;

namespace Red.Core
{
    public static class StringHelper
    {
        /// <summary>
        /// Returns a string equal to the given when when the "CompareWords" comparer is used.
        /// The returned string is trimmed, of a single case, and has only single spaces as whitespace.
        /// </summary>
        public static string GetWords(string a)
        {
            a = Regex
                .Replace(a, @"\s+", " ")
                .Trim()
                .ToUpper();

            return a;
        }

        /// <summary>
        /// Does not consider whitespace or case when comparing. Other symbols (such as hyphens,
        /// or accents) still make a difference, however.
        /// </summary>
        public static bool CompareWords(string a, string b)
        {
            a = GetWords(a);
            b = GetWords(b);

            return a == b;
        }

        public static IEnumerable<string> Split(string source)
        {
             return source
                .Split(new string[] {"\n", "\r", ",", ".", ";"}, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();
        }
    }
}
