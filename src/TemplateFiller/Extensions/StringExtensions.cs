using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TemplateFiller.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// 尝试替换<paramref name="target"/>中，第一个被<paramref name="pattern"/>匹配的字符串为<paramref name="replacement"/>
        /// </summary>
        /// <param name="target">目标字符串</param>
        /// <param name="pattern">匹配模式，应为正则表达式</param>
        /// <param name="replacement">替换文本</param>
        /// <returns>与模式匹配时，返回替换第一个匹配项后的新字符串；否则返回原本的字符串</returns>
        public static string TryReplaceFirstMatch(this string target, string pattern, string replacement)
        {
            var regex = new Regex(pattern);
            var match = regex.Match(target);
            if (match.Success)
            {
                return target.Substring(0, match.Index) +
                       replacement +
                       target.Substring(match.Index + match.Length);
            }
            return target;
        }

        /// <summary>
        /// 检查<paramref name="target"/>中，是否存在被<paramref name="pattern"/>匹配的字符串。若有匹配项，返回匹配的具体情况
        /// </summary>
        /// <param name="target">目标字符串</param>
        /// <param name="pattern">匹配模式，应为正则表达式</param>
        /// <param name="patternOnly">为<c>true</c>时。代表目标字符串完全与模式匹配，不包含其他字符</param>
        /// <param name="matchCount">非0时，代表目标字符串中，与模式匹配的次数</param>
        /// <returns>存在匹配的字符串时，返回<c>true</c>，其他情况下返回<c>false</c></returns>
        public static bool IsMatch(this string target, string pattern, out bool patternOnly, out int matchCount)
        {
            var isMatch = Regex.IsMatch(target, pattern);
            if (!isMatch)
            {
                patternOnly = false;
                matchCount = 0;
                return false;
            }

            var matches = Regex.Matches(target, pattern);
            var matchedText = string.Join("", matches.Select(m => m.Value));
            patternOnly = target.Length == matchedText.Length;
            matchCount = matches.Count;
            return true;
        }

        /// <summary>
        /// 用于解决字符串与模式<paramref name="pattern"/>匹配的部分，被拆分成多个子串时，如何确定哪些子串包含了匹配的字符，以及这些字符的位置
        /// </summary>
        /// <param name="substrings">被拆分的子串</param>
        /// <param name="pattern">有且只有一个分组的模式，比如@"\{(.+?)\}"</param>
        /// <returns>反应匹配情况的字典，键是模式中的分组值，值是数组，反应了匹配的子串索引和子串内的位置和长度</returns>
        /// 
        /// <remarks>
        /// <example>
        /// <code>
        /// var srcString = "起始时间{StartTime} 终止时间{EndTime}";
        /// var subStrings = { "起始时间", "{", "StartTime", "}", " 终止时间", "{", "EndTime", "}" };
        /// var pattern = @"\{(.+?)\}";
        /// var result = subStrings.FindMatchingSubstrings(pattern); // StartTime: 1, 0, 1
        ///                                                          //            2. 0. 9 
        ///                                                          //            3. 0. 1
        ///                                                          // EndTime:   5, 0, 1
        ///                                                          //            6. 0. 7 
        ///                                                          //            7. 0. 1
        /// </code>
        /// </example>
        /// </remarks>

        public static Dictionary<string, List<(int SubstringIndex, int StartIndex, int Length)>> FindMatchingSubstrings(
            this string[] substrings, string pattern)
        {
            // 合并所有子串以重建原始字符串
            var fullString = string.Concat(substrings);

            // 使用正则表达式找到所有匹配的部分
            var matches = Regex.Matches(fullString, pattern);
            if (matches.Count == 0)
                return [];

            var result = new Dictionary<string, List<(int, int, int)>>();
            int currentPos = 0; // 当前在合并字符串中的位置
            int substringIndex = 0; // 当前处理的子串索引

            foreach (Match match in matches)
            {
                string matchedContent = match.Groups[1].Value; // 获取匹配的内容
                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length;

                var positions = new List<(int, int, int)>();

                // 遍历子串，直到覆盖当前匹配项
                while (substringIndex < substrings.Length && currentPos < matchEnd)
                {
                    string sub = substrings[substringIndex];
                    int subLength = sub.Length;
                    int subStart = currentPos;
                    int subEnd = currentPos + subLength;

                    // 检查当前子串是否与匹配项有重叠
                    if (subEnd > matchStart && subStart < matchEnd)
                    {
                        // 计算重叠部分的起始和结束
                        int overlapStart = Math.Max(matchStart, subStart);
                        int overlapEnd = Math.Min(matchEnd, subEnd);
                        int overlapLength = overlapEnd - overlapStart;

                        if (overlapLength > 0)
                        {
                            // 计算在子串中的起始位置
                            int startInSub = overlapStart - subStart;
                            positions.Add((substringIndex, startInSub, overlapLength));
                        }
                    }

                    currentPos = subEnd;
                    substringIndex++;
                }

                // 确保下一个匹配项从正确的子串开始
                substringIndex--;
                currentPos -= substrings[substringIndex].Length;

                result[matchedContent] = positions;
            }

            return result;
        }
    }
}
