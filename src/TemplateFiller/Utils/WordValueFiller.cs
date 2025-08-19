using NPOI.XWPF.UserModel;
using System;
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
    /// <remarks>
    /// 占位符参见：<seealso cref="PlaceholderConsts.ValuePlaceholder"/>
    /// </remarks>
    public sealed class WordValueFiller(XWPFParagraph? paragraph) : ITargetFiller, IDisposable
    {
        private XWPFParagraph? _paragraph { get; set; } = paragraph;

        /// <inheritdoc/>
        public bool Check()
        {
            if (_paragraph == null)
            {
                return false;
            }

            return Regex.IsMatch(_paragraph.Text, PlaceholderConsts.ValuePlaceholder);
        }

        /// <inheritdoc/>
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
                    _paragraph.RemoveRun(1);
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
                var runMatch = match.Key;
                var key = runMatch.Groups[1].Value;
                var subStrings = match.Value;
                var replaceStr = source[key]?.ToString() ?? string.Empty;
                switch (subStrings.Count)
                {
                    case 1:
                        subStrings.RemoveMatchedTextInHeadRun(_paragraph);
                        subStrings.AddReplaceValueToHeadRun(_paragraph, replaceStr);
                        break;
                    case 2:
                        // 模式被分割在两个run中，那么移除头部和尾部与模式匹配的文本。然后在头部末尾添加
                        subStrings.RemoveMatchedTextInHeadRun(_paragraph);
                        subStrings.RemoveMatchedTextInTailRun(_paragraph);
                        subStrings.AddReplaceValueToHeadRun(_paragraph, replaceStr);
                        subStrings.RemoveTailRunIfTextIsEmptyOrNull(_paragraph);
                        break;
                    case >= 3:
                        // 模式被分割在多个run中，那么移除头部、尾部与模式匹配的文本，移除头尾之间的所有run。然后在头部末尾添加
                        subStrings.RemoveMatchedTextInHeadRun(_paragraph);
                        subStrings.RemoveMatchedTextInTailRun(_paragraph);
                        subStrings.RemoveRunsBetweenHeadAndTail(_paragraph);
                        subStrings.AddReplaceValueToHeadRun(_paragraph, replaceStr);
                        subStrings.RemoveTailRunIfTextIsEmptyOrNull(_paragraph);
                        break;
                    default:
                        break;
                }
            }
        }

        private static string BuildReplaceText(ISource source, string originalText)
        {
            return Regex.Replace(originalText, PlaceholderConsts.ValuePlaceholder, match =>
            {
                var key = match.Groups[1].Value;
                return source[key]?.ToString() ?? string.Empty;
            });
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
        public void Dispose()
        {
            _paragraph = null;
        }
    }
}
