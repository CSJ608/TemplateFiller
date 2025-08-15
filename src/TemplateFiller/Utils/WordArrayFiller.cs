using NPOI.Util;
using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;
using TemplateFiller.Extensions;

namespace TemplateFiller.Utils
{
    /// <summary>
    /// 表示Word表格填充器
    /// <para>可以识别表格内是否有占位符[array]或者[array.property]</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[array]，然后一边遍历source[array]，一边填充目标单元格。填充一次完成后，
    /// 就将目标向下移动一格。若发现合并的单元格，或者未合并但有数据的单元格，则会插入一行。</para>
    /// <para>取值时，若占位符是[array]，则把遍历时的单个元素item直接写入单元格；若占位符是[array.property]，则取item[property]的值写入单元格。</para>
    /// </summary>
    /// <remarks>
    /// 占位符参见：<seealso cref="PlaceholderConsts.ArrayPlaceholder"/>
    /// </remarks>
    public class WordArrayFiller : ITargetFiller, IDisposable
    {
        private XWPFTable? _table {  get; set; }
        private Dictionary<int, (IEnumerator enumerator, string? propertyName)?> ReplaceSource {  get; set; }

        public WordArrayFiller(XWPFTable? table)
        {
            _table = table;
            ReplaceSource = [];
        }

        /// <inheritdoc/>
        public bool Check()
        {
            if(_table == null)
            {
                return false;
            }

            foreach (var row in _table.Rows) {
                foreach (var cell in row.GetTableCells())
                {
                    var str = cell.GetText();
                    var match = Regex.Match(str, PlaceholderConsts.ArrayPlaceholder);
                    if (match.Success)
                    {
                        return true;
                    }
                }
            }
            
            return false;
        }

        public void Fill(ISource source)
        {
            if(_table == null)
            {
                return;
            }

            var startRowIndex = 0;
            while (startRowIndex < _table.Rows.Count)
            {
                var row = _table.Rows[startRowIndex];
                var cells = row.GetTableCells();
                ResetReplaceSource();
                BuildReplaceSource(source, cells);

                // 一整行内没有占位符，考虑下一行
                if (IsReplaceSourceEmpty())
                {
                    startRowIndex++;
                    continue;
                }

                // 当前行内至少有一列存在占位符。
                // 填充当前行的所有占位符
                var rowOffset = 1; 
                var templateRow = _table.Rows[startRowIndex];
                foreach (var allReplaceValues in GetRowReplaceValues())
                {
                    // 若目标行不存在，则在表格末尾添加一行模板行的拷贝
                    // 若目标行存在，并且待填充的各列内容都为空，且不存在合并的单元格，则删除目标行后，再插入模板行的拷贝
                    // 若目标行存在，并且待填充的各列存在某一列内容不为空，或存在合并的单元格，则插入模板行的拷贝
                    // 最后对插入的模板行的各列进行占位符替换

                    var targetRowIndex = startRowIndex + rowOffset;
                    var targetRow = _table.GetRow(targetRowIndex);
                    if(targetRow == null)
                    {
                        targetRow = templateRow.CloneRow(targetRowIndex);
                    }
                    else
                    {
                        // 检查目标行中是否存在合并的单元格，或者任意文本
                        var hasText = false;
                        var existsMerged = false;
                        foreach (var (columnIndex, _) in allReplaceValues)
                        {
                            var cell = targetRow.GetCell(columnIndex);
                            if (!string.IsNullOrWhiteSpace(cell.GetText()))
                            {
                                hasText = true;
                                break;
                            }

                            if (cell.IsCellMerged())
                            {
                                existsMerged = true;
                                break;
                            }
                        }

                        targetRow = templateRow.CloneRow(targetRowIndex);
                        if (!hasText && !existsMerged)
                        {
                            _table.RemoveRow(targetRowIndex + 1);
                        }
                    }

                    // 填充目标行内待填充的各列                    

                    foreach (var (columnIndex, replaceStr) in allReplaceValues)
                    {
                        var cell = targetRow.GetCell(columnIndex);
                        FillCurrentCell(cell, replaceStr);
                    }
                    rowOffset++;
                }

                // 删除模板行
                _table.RemoveRow(startRowIndex);
                startRowIndex += rowOffset - 1;
            }
        }

        private void BuildReplaceSource(ISource source, List<XWPFTableCell> cells)
        {
            for (int columnIndex = 0; columnIndex < cells.Count; columnIndex++)
            {
                var cell = cells[columnIndex];
                var str = cell.GetText();

                if (!str.IsMatch(PlaceholderConsts.ArrayPlaceholder, out var patternOnly, out var matchCount))
                {
                    continue;
                }

                var match = Regex.Match(str, PlaceholderConsts.ArrayPlaceholder); // 同一个单元格内有多个时，只匹配一个
                if (!match.Success)
                {
                    continue;
                }

                // 需要在当前单元格及其下方填充数据
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
                    continue;
                }

                if (!(section.Value is IEnumerable enumerable))
                {
                    continue;
                }
                var enumerator = enumerable.GetEnumerator();
                AddReplaceSource(columnIndex, enumerator, propertyName);
            }
        }

        private void ResetReplaceSource()
        {
            ReplaceSource = [];
        }

        private bool IsReplaceSourceEmpty()
        {
            return ReplaceSource.Keys.Count == 0;
        }

        private void AddReplaceSource(int columnIndex, IEnumerator enumerator, string? propertyName)
        {
            if (!ReplaceSource.ContainsKey(columnIndex)) {
                ReplaceSource.Add(columnIndex, (enumerator, propertyName));
            }
        }

        private IEnumerable<IEnumerable<(int columnIndex, string replaceStr)>> GetRowReplaceValues()
        {
            var columnIndexs = ReplaceSource.Keys;
            var emptyIndexs = new List<int>();
            while( emptyIndexs.Count < columnIndexs.Count)
            {
                var result = new List<(int columnIndex, string replaceStr)>();
                foreach (var columnIndex in columnIndexs)
                {
                    string replaceStr = string.Empty;
                    var source = ReplaceSource[columnIndex];
                    if (source.HasValue)
                    {
                        var enumerator = source.Value.enumerator;
                        if (enumerator.MoveNext())
                        {
                            if(source.Value.propertyName == null)
                            {
                                replaceStr = enumerator.Current?.ToString() ?? string.Empty;
                            }
                            else
                            {
                                using var s = new Source(enumerator.Current);
                                replaceStr = s[source.Value.propertyName]?.ToString() ?? string.Empty;
                            }
                        }
                        else
                        {
                            emptyIndexs.Add(columnIndex);
                            ReplaceSource[columnIndex] = null;
                        }
                    }

                    result.Add((columnIndex, replaceStr));
                }

                if (emptyIndexs.Count < columnIndexs.Count) {
                    // 还存在枚举器没遍历完，
                    yield return result;
                }
            }
            
        }

        private void FillCurrentCell(XWPFTableCell currentCell, string replaceStr)
        {
            foreach (var p in currentCell.Paragraphs)
            {
                var matches = p.Runs.Select(t => t.Text).ToArray().FindMatchingSubstrings(PlaceholderConsts.ArrayPlaceholder);
                if (matches.Count == 0)
                {
                    continue;
                }

                // 对于数组元素只替换第一个
                var match = matches.First();
                var runMatch = match.Key;
                var key = runMatch.Groups[1].Value;
                var subStrings = match.Value;
                switch (subStrings.Count)
                {
                    case 1:
                        subStrings.RemoveMatchedTextInHeadRun(p);
                        subStrings.AddReplaceValueToHeadRun(p, replaceStr);
                        break;
                    case 2:
                        // 模式被分割在两个run中，那么移除头部和尾部与模式匹配的文本。然后在头部末尾添加
                        subStrings.RemoveMatchedTextInHeadRun(p);
                        subStrings.RemoveMatchedTextInTailRun(p);
                        subStrings.AddReplaceValueToHeadRun(p, replaceStr);
                        subStrings.RemoveTailRunIfTextIsEmptyOrNull(p);
                        break;
                    case >= 3:
                        // 模式被分割在多个run中，那么移除头部、尾部与模式匹配的文本，移除头尾之间的所有run。然后在头部末尾添加
                        subStrings.RemoveMatchedTextInHeadRun(p);
                        subStrings.RemoveMatchedTextInTailRun(p);
                        subStrings.RemoveRunsBetweenHeadAndTail(p);
                        subStrings.AddReplaceValueToHeadRun(p, replaceStr);
                        subStrings.RemoveTailRunIfTextIsEmptyOrNull(p);
                        break;
                    default:
                        break;
                }

                break; // 一个单元格内如果有多个占位符，只处理匹配的第一个 
            }
        }

        /// <summary>
        /// 更换目标
        /// </summary>
        /// <param name="table"></param>
        public void ChangeTarget(XWPFTable table)
        {
            _table = table;
        }

        public void Dispose()
        {
            _table = null;
        }
    }
}
