using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TemplateFiller.Extensions
{
    public static class XWPFExtensions
    {
        /// <summary>
        /// 用于解决字符串与模式<paramref name="pattern"/>匹配的部分，被拆分成多个子串时，如何确定哪些子串包含了匹配的字符，以及这些字符的位置
        /// </summary>
        /// <param name="substrings">被拆分的子串</param>
        /// <param name="pattern">匹配模式</param>
        /// <returns>反应匹配情况的字典，键是匹配的结果，值是数组，反应了匹配的子串索引和子串内的位置和长度</returns>
        /// 
        /// <remarks>
        /// <example>
        /// <code>
        /// var srcString = "起始时间{StartTime} 终止时间{EndTime}";
        /// var subStrings = { "起始时间", "{", "StartTime", "}", " 终止时间", "{", "EndTime", "}" };
        /// var pattern = @"\{(.+?)\}";
        /// var result = subStrings.FindMatchingSubstrings(pattern); // {StartTime}: 1, 0, 1
        ///                                                          //              2. 0. 9 
        ///                                                          //              3. 0. 1
        ///                                                          // {EndTime}:   5, 0, 1
        ///                                                          //              6. 0. 7 
        ///                                                          //              7. 0. 1
        /// </code>
        /// </example>
        /// </remarks>

        public static Dictionary<Match, List<(int SubstringIndex, int StartIndex, int Length)>> FindMatchingSubstrings(
            this string[] substrings, string pattern)
        {
            // 合并所有子串以重建原始字符串
            var fullString = string.Concat(substrings);

            // 使用正则表达式找到所有匹配的部分
            var matches = Regex.Matches(fullString, pattern);
            if (matches.Count == 0)
                return [];

            var result = new Dictionary<Match, List<(int, int, int)>>();
            int currentPos = 0; // 当前在合并字符串中的位置
            int substringIndex = 0; // 当前处理的子串索引

            foreach (Match match in matches)
            {
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

                result[match] = positions;
            }

            return result;
        }

        public static void RemoveTailRunIfTextIsEmptyOrNull(this List<(int SubstringIndex, int StartIndex, int Length)> subStrings, XWPFParagraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            var headerIndex = subStrings.First().SubstringIndex;
            var tailIndex = headerIndex + 1;
            if (string.IsNullOrEmpty(paragraph.Runs[tailIndex].Text))
            {
                paragraph.RemoveRun(tailIndex);
            }
        }

        public static void RemoveRunsBetweenHeadAndTail(this List<(int SubstringIndex, int StartIndex, int Length)> subStrings, XWPFParagraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            var tailInfo = subStrings.Last();
            var headInfo = subStrings.First();
            for (int i = tailInfo.SubstringIndex - 1; i > headInfo.SubstringIndex; i--)
            {
                paragraph.RemoveRun(i);
            }
        }

        public static void RemoveMatchedTextInHeadRun(this List<(int SubstringIndex, int StartIndex, int Length)> subStrings, XWPFParagraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            var headInfo = subStrings.First();
            var headRun = paragraph.Runs[headInfo.SubstringIndex];            
            var newText = headRun.Text.RemoveAt(headInfo.StartIndex, headInfo.Length);            
            headRun.SetText(newText);
        }

        public static void RemoveMatchedTextInTailRun(this List<(int SubstringIndex, int StartIndex, int Length)> subStrings, XWPFParagraph paragraph)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            var tailInfo = subStrings.Last();
            var tailRun = paragraph.Runs[tailInfo.SubstringIndex];
            var newText = tailRun.Text.RemoveAt(tailInfo.StartIndex, tailInfo.Length);
            tailRun.SetText(newText);
        }

        public static void AddReplaceValueToHeadRun(this List<(int SubstringIndex, int StartIndex, int Length)> subStrings, XWPFParagraph paragraph, string replaceValue)
        {
            if (paragraph == null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            var headInfo = subStrings.First();
            var headRun = paragraph.Runs[headInfo.SubstringIndex];
            var newText = headRun.Text.InsertAt(headInfo.StartIndex, replaceValue);
            headRun.SetText(newText);
        }
    }
}
