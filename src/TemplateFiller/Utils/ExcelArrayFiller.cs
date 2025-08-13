using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Text.RegularExpressions;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;
using TemplateFiller.Extensions;

namespace TemplateFiller.Utils
{
    /// <summary>
    /// 表示Excel数组填充器。
    /// <para>可以识别表格内是否有占位符[array:property]</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[array]，然后尝试遍历其中每个元素item，依次获取item[property]的值， 然后替换占位符。
    /// 每替换一次，活动单元格就向下移动一次。在向下移动的过程中，若发现合并的单元格，或者未合并但有数据的单元格，则插入一行，再填写数据</para>
    /// </summary>
    public class ExcelArrayFiller : IFiller, IDisposable
    {
        private ICell? _cell { get; set; }

        public ExcelArrayFiller(ICell? cell)
        {
            _cell = cell;
        }

        /// <inheritdoc/>
        public bool Check()
        {
            if (_cell == null)
            {
                return false;
            }

            var match = Regex.Match(_cell.GetStringValue(), PlaceholderConsts.ArrayPlaceholder);
            return match.Success;
        }

        public void Dispose()
        {
            _cell = null;
        }

        /// <inheritdoc/>
        public void Fill(ISource source)
        {
            if (_cell == null)
            {
                return;
            }

            var str = _cell.GetStringValue();
            if (!str.IsMatch(PlaceholderConsts.ArrayPlaceholder, out var patternOnly, out var matchCount))
            {
                return;
            }

            var match = Regex.Match(str, PlaceholderConsts.ArrayPlaceholder); // 同一个单元格内有多个时，只匹配一个
            
            var arrayPath = match.Groups["collectionPath"].Value;
            var fullMatch = match.Value;
            string? propertyName = null;
            if (match.Groups["propertyPath"].Success)
            {
                propertyName = match.Groups["propertyPath"].Value;
            }

            // 获取数组数据
            var section = source.GetSection(arrayPath);
            if (section == null)
            {
                return;
            }

            if (!(section.Value is IEnumerable enumerable))
            {
                return;
            }

            var sheet = _cell.Sheet;
            var startRow = _cell.RowIndex;
            var column = _cell.ColumnIndex;
            var rowOffset = 0;
            var templateRow = _cell.Row;
            foreach (var item in enumerable)
            {
                using var s = new Source(item);
                var value = propertyName == null ? item?.ToString() ?? string.Empty : s[propertyName];
                var replaceStr = str.TryReplaceFirstMatch(PlaceholderConsts.ArrayPlaceholder, value?.ToString() ?? string.Empty);
                var filledValue = patternOnly && matchCount == 1 ? value : replaceStr;

                var rowIndex = startRow + rowOffset;
                var currentRow = sheet.GetRow(rowIndex) ?? sheet.CreateRowAndCopyStyle(rowIndex, templateRow);
                var currentCell = currentRow.GetCell(column);

                if (currentCell == null)
                {
                    currentCell = currentRow.CreateCell(column);
                    currentCell.CopyStyle(templateRow);
                }
                else if (currentCell.IsMergedCell || !string.IsNullOrEmpty(currentCell.StringCellValue) && rowOffset > 0)
                {
                    // 创建一行
                    sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1);
                    currentRow = sheet.CreateRowAndCopyStyle(rowIndex, templateRow);
                    currentCell = currentRow.GetCell(column);
                }
                else
                {
                    currentCell.CopyStyle(templateRow);
                }

                currentCell.SetExcelCellValueByType(filledValue);
                rowOffset++;
            }
        }

        /// <summary>
        /// 更换目标
        /// </summary>
        /// <param name="cell"></param>
        public void ChangeTarget(ICell? cell)
        {
            _cell = cell;
        }
    }
}
