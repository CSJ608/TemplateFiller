using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;
using TemplateFiller.Extensions;

namespace TemplateFiller.Utils
{
    /// <summary>
    /// 表示Word数据填充器。
    /// <para>可以识别段落内是否有占位符{path}</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[path]，然后替换占位符</para>
    /// </summary>
    public sealed class WordValueFiller : IFiller, IDisposable
    {
        private XWPFParagraph? _paragraph { get; set; }

        public WordValueFiller(XWPFParagraph? paragraph)
        {
            _paragraph = paragraph;
        }

        /// <summary>
        /// 更换目标
        /// </summary>
        /// <param name="paragraph"></param>
        public void ChangeTarget(XWPFParagraph? paragraph)
        {
            _paragraph = paragraph;
        }

        /// <inheritdoc/>
        public bool Check()
        {
            if (_paragraph == null)
            {
                return false;
            }

            return Regex.IsMatch(_paragraph.Text, PlaceholderConsts.ValuePlaceholder);
        }

        public void Fill(ISource source)
        {
            if (_paragraph == null)
            {
                return;
            }

            if (!_paragraph.Text.IsMatch(PlaceholderConsts.ValuePlaceholder, out var patternOnly, out var matchCount))
            {
                return;
            }

            // todo 移除run前需要考虑run内是否有图片。如果有图片的话就不能移除run
            if (patternOnly)
            {
                var replaceStr = BuildReplaceText(source, _paragraph.Text);
                while (_paragraph.Runs.Count > 1)
                {
                    _paragraph.Runs.RemoveAt(1);
                }
                _paragraph.Runs[0].SetText(replaceStr);
                return;
            }

            var matches = _paragraph.Runs.Select(t => t.Text).ToArray().FindMatchingSubstrings(PlaceholderConsts.ValuePlaceholder);
            if (matches.Count == 0)
            {
                return;
            }

            // 倒序处理。否则插入、删除run会造成后续run的索引被改变
            foreach (var match in matches.Reverse())
            {
                var key = match.Key;
                var subStrings = match.Value;
                switch (subStrings.Count)
                {
                    case 1:
                        RemoveMatchedTextInHeadRun(subStrings);
                        AddReplaceValueToHeadRun(source, key, subStrings);
                        break;
                    case 2:
                        // 模式被分割在两个run中，那么移除头部和尾部与模式匹配的文本。然后在头部末尾添加
                        RemoveMatchedTextInHeadRun(subStrings);
                        RemoveMatchedTextInTailRun(subStrings);
                        AddReplaceValueToHeadRun(source, key, subStrings);
                        RemoveTailRunIfTextIsEmptyOrNull(subStrings);
                        break;
                    case >= 3:
                        // 模式被分割在多个run中，那么移除头部、尾部与模式匹配的文本，移除头尾之间的所有run。然后在头部末尾添加
                        RemoveMatchedTextInHeadRun(subStrings);
                        RemoveMatchedTextInTailRun(subStrings);
                        RemoveRunsBetweenHeadAndTail(subStrings);
                        AddReplaceValueToHeadRun(source, key, subStrings);
                        RemoveTailRunIfTextIsEmptyOrNull(subStrings);
                        break;
                    default:
                        break;
                }
            }
        }

        private void RemoveTailRunIfTextIsEmptyOrNull(List<(int SubstringIndex, int StartIndex, int Length)> subStrings)
        {
            if (_paragraph == null)
            {
                throw new ArgumentNullException(nameof(_paragraph));
            }

            var headerIndex = subStrings.First().SubstringIndex;
            var tailIndex = headerIndex + 1;
            if (string.IsNullOrEmpty(_paragraph.Runs[tailIndex].Text))
            {
                _paragraph.RemoveRun(tailIndex);
            }
        }

        private void RemoveRunsBetweenHeadAndTail(List<(int SubstringIndex, int StartIndex, int Length)> subStrings)
        {
            if (_paragraph == null)
            {
                throw new ArgumentNullException(nameof(_paragraph));
            }

            var tailInfo = subStrings.Last();
            var headInfo = subStrings.First();
            for (int i = tailInfo.SubstringIndex - 1; i > headInfo.SubstringIndex; i--)
            {
                _paragraph.RemoveRun(i);
            }
        }

        private void AddReplaceValueToHeadRun(ISource source, string key, List<(int SubstringIndex, int StartIndex, int Length)> subStrings)
        {
            if (_paragraph == null)
            {
                throw new ArgumentNullException(nameof(_paragraph));
            }

            var value = source[key]?.ToString() ?? string.Empty;
            var headInfo = subStrings.First();
            var headRun = _paragraph.Runs[headInfo.SubstringIndex];
            headRun.SetText(headRun.Text + value);
        }

        private void RemoveMatchedTextInHeadRun(List<(int SubstringIndex, int StartIndex, int Length)> subStrings)
        {
            if (_paragraph == null)
            {
                throw new ArgumentNullException(nameof(_paragraph));
            }

            var headInfo = subStrings.First();
            var headRun = _paragraph.Runs[headInfo.SubstringIndex];
            var textLengthWithoutPlaceholder = headRun.Text.Length - headInfo.Length;
            headRun.SetText(headRun.Text.Substring(0, textLengthWithoutPlaceholder));
        }

        private void RemoveMatchedTextInTailRun(List<(int SubstringIndex, int StartIndex, int Length)> subStrings)
        {
            if (_paragraph == null)
            {
                throw new ArgumentNullException(nameof(_paragraph));
            }

            var tailInfo = subStrings.Last();
            var tailRun = _paragraph.Runs[tailInfo.SubstringIndex];
            var textLengthWithoutPlaceholder = tailRun.Text.Length - tailInfo.Length;
            tailRun.SetText(tailRun.Text.Substring(tailInfo.Length, textLengthWithoutPlaceholder));
        }

        private static string BuildReplaceText(ISource source, string originalText)
        {
            return Regex.Replace(originalText, PlaceholderConsts.ValuePlaceholder, match =>
            {
                var key = match.Groups[1].Value;
                return source[key]?.ToString() ?? string.Empty;
            });
        }

        public void Dispose()
        {
            _paragraph = null;
        }
    }
}
