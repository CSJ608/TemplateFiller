using NPOI.SS.UserModel;
using System;
using System.Text.RegularExpressions;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;
using TemplateFiller.Extensions;

namespace TemplateFiller.Utils
{
    /// <summary>
    /// 表示Excel数据填充器。
    /// <para>可以识别单元格内是否有占位符{path}</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[path]，然后替换占位符</para>
    /// </summary>
    /// <remarks>
    /// 占位符参见：<seealso cref="PlaceholderConsts.ValuePlaceholder"/>
    /// </remarks>
    public sealed class ExcelValueFiller(ICell? cell) : ITargetFiller, IDisposable
    {
        private ICell? _cell { get; set; } = cell;

        /// <inheritdoc/>
        public bool Check()
        {
            if (_cell == null)
            {
                return false;
            }

            return Regex.IsMatch(_cell.GetStringValue(), PlaceholderConsts.ValuePlaceholder);
        }

        /// <inheritdoc/>
        public void Fill(ISource source)
        {
            if (_cell == null)
            {
                return;
            }

            var str = _cell.GetStringValue();

            if (!str.IsMatch(PlaceholderConsts.ValuePlaceholder, out var patternOnly, out var matchCount))
            {
                return;
            }

            object? value = null;
            var replaceStr = Regex.Replace(str, PlaceholderConsts.ValuePlaceholder, match =>
            {
                var key = match.Groups[1].Value;
                value = source[key];
                return value?.ToString() ?? string.Empty;
            });

            var filledValue = patternOnly && matchCount == 1 ? value : replaceStr;
            _cell.SetExcelCellValueByType(filledValue);
        }

        /// <summary>
        /// 更换目标
        /// </summary>
        /// <param name="cell"></param>
        public void ChangeTarget(ICell? cell)
        {
            _cell = cell;
        }


        /// <inheritdoc/>
        public void Dispose()
        {
            _cell = null;
        }
    }
}
