using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;

namespace TemplateFiller.Extensions
{
    internal static class NPOIExtensions
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

        public static bool IsCellMerged(this XWPFTableCell cell)
        {
            if (cell == null) return false;

            var tcPr = cell.GetCTTc().tcPr;
            if (tcPr == null) return false;

            // 检查横向合并
            if (tcPr.gridSpan != null && int.Parse(tcPr.gridSpan.val) > 1)
                return true;

            // 检查水平合并
            if (tcPr.hMerge != null && (tcPr.hMerge.val == ST_Merge.restart || tcPr.hMerge.val == ST_Merge.@continue))
                return true;

            // 检查垂直合并
            if (tcPr.vMerge != null && (tcPr.vMerge.val == ST_Merge.restart || tcPr.vMerge.val == ST_Merge.@continue))
                return true;

            return false;
        }
    }
}
