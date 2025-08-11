using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using TemplateFiller.Abstractions;

namespace TemplateFiller.Extensions
{
    /// <summary>
    /// 表示Excel数据填充器。
    /// <para>可以识别单元格内是否有占位符{path}</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[path]，然后替换占位符</para>
    /// </summary>
    public sealed class ExcelValueFiller : IFiller, IDisposable
    {
        /// <summary>
        /// 占位符
        /// </summary>
        /// <remarks>
        /// <para>示例</para>
        /// <para>{PrintTime}</para>
        /// <para>{User:Name}</para>
        /// </remarks>
        public const string Placeholder = @"\{(.+?)\}";

        private ICell? _cell { get; set; }
        
        public ExcelValueFiller(ICell? target) 
        {
            _cell = target;
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
        public bool Check()
        {
            if(_cell == null)
            {
                return false;
            }

            return Regex.IsMatch(_cell.GetStringValue(), Placeholder);
        }

        /// <inheritdoc/>
        public void Fill(ISource source)
        {
            if(_cell == null)
            {
                return;
            }

            var str = _cell.GetStringValue();

            if (!Utils.IsMatch(str, Placeholder, out var patternOnly, out var matchCount))
            {
                return;
            }

            object? value = null;
            var replaceStr = Regex.Replace(str, Placeholder, match =>
            {
                var key = match.Groups[1].Value;
                value = source[key];
                return value?.ToString() ?? string.Empty;
            });

            var filledValue = patternOnly && matchCount == 1 ? value : replaceStr;
            _cell.SetExcelCellValueByType(filledValue);
        }

        public void Dispose()
        {
            _cell = null;
        }
    }
}
