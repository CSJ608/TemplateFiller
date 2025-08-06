using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace TemplateFiller
{
    public class NpoiExcelTemplateFiller
    {
        public void FillTemplate(string templatePath, string outputPath, object dataSource)
        {
            using (var provider = new DefaultDataProvider(dataSource))
            {
                IWorkbook workbook = LoadWorkbook(templatePath);

                // 处理所有工作表
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    ProcessSheet(sheet, provider);
                }

                SaveWorkbook(workbook, outputPath);
            }
        }

        private IWorkbook LoadWorkbook(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return Path.GetExtension(filePath).ToLower() == ".xls"
                    ? new HSSFWorkbook(stream) as IWorkbook
                    : new XSSFWorkbook(stream) as IWorkbook;
            }
        }

        private void SaveWorkbook(IWorkbook workbook, string outputPath)
        {
            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }

        private void ProcessSheet(ISheet sheet, IDataProvider provider)
        {
            // 先处理可能包含数组占位符的单元格
            ProcessArrayPlaceholders(sheet, provider);

            // 然后处理普通占位符
            ProcessRegularPlaceholders(sheet, provider);
        }

        private void ProcessArrayPlaceholders(ISheet sheet, IDataProvider provider)
        {
            for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var row = sheet.GetRow(rowNum);
                if (row == null) continue;

                for (int colNum = row.FirstCellNum; colNum < row.LastCellNum; colNum++)
                {
                    var cell = row.GetCell(colNum);
                    if (cell?.CellType != CellType.String) continue;

                    var cellValue = cell.StringCellValue;
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // 检查是否是数组占位符格式 [ArrayName:Property]
                    var arrayMatch = Regex.Match(cellValue, @"^\[(.+?):(.+?)\]$");
                    if (arrayMatch.Success)
                    {
                        var arrayPath = arrayMatch.Groups[1].Value;
                        var propertyName = arrayMatch.Groups[2].Value;

                        // 获取数组数据
                        var items = provider.GetChildren(arrayPath).ToList();
                        if (items.Count > 0)
                        {
                            FillArrayData(sheet, rowNum, colNum, items, propertyName);
                        }
                    }
                }
            }
        }

        private IRow CreateRowAndCopyStyle(ISheet sheet, int rowIndex, IRow templateRow)
        {
            var row = sheet.CreateRow(rowIndex);
            // 复制样式
            for (int col = templateRow.FirstCellNum; col < templateRow.LastCellNum; col++)
            {
                var srcCell = templateRow.GetCell(col);
                if (srcCell != null)
                {
                    var newCell = row.CreateCell(col);
                    newCell.CellStyle = srcCell.CellStyle;
                }
            }

            return row;
        }

        private void FillArrayData(ISheet sheet, int startRow, int column, List<IDataProvider> items, string propertyName)
        {
            var templateRow = sheet.GetRow(startRow);
            for (int i = 0; i < items.Count; i++)
            {
                var rowIndex = startRow + i;
                var currentRow = sheet.GetRow(rowIndex) ?? CreateRowAndCopyStyle(sheet, rowIndex, templateRow);
                var currentCell = currentRow.GetCell(column);

                if(currentCell == null)
                {
                    currentCell = currentRow.CreateCell(column);
                }
                else if (!string.IsNullOrEmpty(currentCell.StringCellValue) && i > 0)
                {
                    // 创建一行
                    sheet.ShiftRows(rowIndex, sheet.LastRowNum, 1);
                    currentRow = CreateRowAndCopyStyle(sheet, rowIndex, templateRow);
                    currentCell = currentRow.GetCell(column);
                }

                var value = items[i][propertyName]?.ToString() ?? "";
                currentCell.SetCellValue(value);
            }
        }

        private void ProcessRegularPlaceholders(ISheet sheet, IDataProvider provider)
        {
            for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var row = sheet.GetRow(rowNum);
                if (row == null) continue;

                for (int colNum = row.FirstCellNum; colNum < row.LastCellNum; colNum++)
                {
                    var cell = row.GetCell(colNum);
                    if (cell?.CellType != CellType.String) continue;

                    var cellValue = cell.StringCellValue;
                    if (string.IsNullOrEmpty(cellValue)) continue;

                    // 跳过已经被处理的数组占位符
                    if (Regex.IsMatch(cellValue, @"^\[.+?:.+?\]$")) continue;

                    // 处理普通占位符 {PropertyPath}
                    if (Regex.IsMatch(cellValue, @"\{.+?\}"))
                    {
                        var newValue = Regex.Replace(cellValue, @"\{(.+?)\}", match =>
                        {
                            var key = match.Groups[1].Value;
                            return provider[key]?.ToString() ?? "";
                        });

                        cell.SetCellValue(newValue);
                    }
                }
            }
        }
    }
}
