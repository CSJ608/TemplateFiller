using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TemplateFiller.Extensions
{
    public static class Utils
    {
        public static string ReplaceFirstMatch(string input, string pattern, string replacement)
        {
            Regex regex = new Regex(pattern);
            Match match = regex.Match(input);
            if (match.Success)
            {
                return input.Substring(0, match.Index) +
                       replacement +
                       input.Substring(match.Index + match.Length);
            }
            return input;
        }

        public static bool IsMatch(string input, string pattern, out bool patternOnly, out int matchCount) 
        {
            var isMatch = Regex.IsMatch(input, pattern);
            if (!isMatch)
            {
                patternOnly = false;
                matchCount = 0;
                return false;
            }

            var matches = Regex.Matches(input, pattern);
            var matchedText = string.Join("", matches.Select(m => m.Value));
            patternOnly = input.Length == matchedText.Length;
            matchCount = matches.Count;
            return true;
        }
    }
}
