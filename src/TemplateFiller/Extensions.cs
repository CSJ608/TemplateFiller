using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller
{
    public static class NPOIExtensions
    {
        public static bool HasData(this ISheet sheet, int rowNum)
        {
            var row = sheet.GetRow(rowNum);
            if (row == null) return false;

            foreach (var cell in row.Cells)
            {
                if (cell != null && !string.IsNullOrEmpty(cell.ToString()))
                    return true;
            }
            return false;
        }

        public static void CopyRowStyle(this IRow sourceRow, IRow targetRow)
        {
            foreach (var srcCell in sourceRow.Cells)
            {
                var newCell = targetRow.GetCell(srcCell.ColumnIndex) ?? targetRow.CreateCell(srcCell.ColumnIndex);
                newCell.CellStyle = srcCell.CellStyle;
            }
        }
    }
}
