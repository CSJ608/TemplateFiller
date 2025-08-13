using NPOI.SS.UserModel;
using System;

namespace TemplateFiller.Extensions
{
    public static class NPOIExtensions
    {
        public static void SetExcelCellValueByType(this NPOI.SS.UserModel.ICell cell, object? value)
        {
            if (value == null)
            {
                cell.SetBlank();
                return;
            }

            switch (value)
            {
                case string s:
                    cell.SetCellValue(s);
                    break;
                case bool b:
                    cell.SetCellValue(b);
                    break;
                case DateTime dt:
                    cell.SetCellValue(dt);
                    break;
                case double d:
                    cell.SetCellValue(d);
                    break;
                case int i:
                    cell.SetCellValue(i);
                    break;
                case decimal dec:
                    cell.SetCellValue((double)dec);
                    break;
                case float f:
                    cell.SetCellValue(f);
                    break;

                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }

        public static IRow CreateRowAndCopyStyle(this ISheet sheet, int rowIndex, IRow templateRow)
        {
            var row = sheet.CreateRow(rowIndex);
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

        public static bool IsCellMerged(this NPOI.SS.UserModel.ICell cell)
        {
            var rowIndex = cell.RowIndex;
            var columnIndex = cell.ColumnIndex;
            foreach (var region in cell.Sheet.MergedRegions)
            {
                if (rowIndex >= region.FirstRow && rowIndex <= region.LastRow
                    && columnIndex >= region.FirstColumn && columnIndex <= region.LastColumn)
                {
                    return true;
                }
            }

            return false;
        }

        public static void CopyStyle(this NPOI.SS.UserModel.ICell cell, IRow templateRow)
        {
            var srcCell = templateRow.GetCell(cell.ColumnIndex);
            cell.CellStyle = srcCell.CellStyle;
        }

        public static string GetStringValue(this NPOI.SS.UserModel.ICell cell)
        {
            var oldType = cell.CellType;
            cell.SetCellType(CellType.String);
            var str = cell.StringCellValue;
            cell.SetCellType(oldType);
            return str;
        }
    }
}
