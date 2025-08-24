using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace TemplateFiller.Extensions
{
    internal static class StringExtensions
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
                return target[..match.Index] +
                       replacement +
                       target[(match.Index + match.Length)..];
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
        /// 从字符串中移除指定位置开始的指定长度字符
        /// </summary>
        /// <param name="str">原始字符串</param>
        /// <param name="startIndex">开始移除的位置(从0开始)</param>
        /// <param name="length">要移除的长度</param>
        /// <returns>处理后的新字符串</returns>
        public static string RemoveAt(this string str, int startIndex, int length)
        {
            if (string.IsNullOrEmpty(str))
                return str;

            if (startIndex < 0 || startIndex >= str.Length)
                throw new ArgumentOutOfRangeException(nameof(startIndex), "起始位置超出字符串范围");

            if (length < 0 || startIndex + length > str.Length)
                throw new ArgumentOutOfRangeException(nameof(length), "要移除的长度超出字符串范围");

            return str[..startIndex] + str[(startIndex + length)..];
        }

        /// <summary>
        /// 在字符串的指定位置插入另一个字符串
        /// </summary>
        /// <param name="str">原始字符串</param>
        /// <param name="startIndex">插入位置(从0开始)</param>
        /// <param name="value">要插入的字符串</param>
        /// <returns>处理后的新字符串</returns>
        public static string InsertAt(this string str, int startIndex, string value)
        {
            if (string.IsNullOrEmpty(str))
                return value ?? string.Empty;

            if (startIndex < 0 || startIndex > str.Length)
                throw new ArgumentOutOfRangeException(nameof(startIndex), "插入位置超出字符串范围");

            if (value == null)
                return str;

            return str[..startIndex] + value + str[startIndex..];
        }
    }
}
