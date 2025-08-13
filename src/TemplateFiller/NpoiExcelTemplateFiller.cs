using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using TemplateFiller.Abstractions;
using TemplateFiller.Utils;

namespace TemplateFiller
{
    public class NpoiExcelTemplateFiller
    {
        public static bool CreateDirectoryIfNotExists { get; set; } = true;
        public void FillTemplate(string templatePath, string outputPath, object dataSource)
        {
            using (var source = new Source(dataSource))
            {
                IWorkbook workbook = LoadWorkbook(templatePath);

                // 处理所有工作表
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    ProcessSheet(sheet, source);
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
            if (CreateDirectoryIfNotExists)
            {
                var directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }

            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }

        private void ProcessSheet(ISheet sheet, ISource source)
        {
            // 先处理可能包含数组占位符的单元格
            ProcessArrayPlaceholders(sheet, source);

            // 然后处理普通占位符
            ProcessRegularPlaceholders(sheet, source);
        }

        private void ProcessArrayPlaceholders(ISheet sheet, ISource source)
        {
            var filler = new ExcelArrayFiller(null);
            for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var row = sheet.GetRow(rowNum);
                if (row == null) continue;

                for (int colNum = row.FirstCellNum; colNum < row.LastCellNum; colNum++)
                {
                    var cell = row.GetCell(colNum);
                    filler.ChangeTarget(cell);
                    if (!filler.Check())
                    {
                        continue;
                    }

                    filler.Fill(source);
                }
            }
        }

        private void ProcessRegularPlaceholders(ISheet sheet, ISource source)
        {
            var filler = new ExcelValueFiller(null);
            for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var row = sheet.GetRow(rowNum);
                if (row == null) continue;

                for (int colNum = row.FirstCellNum; colNum < row.LastCellNum; colNum++)
                {
                    var cell = row.GetCell(colNum);
                    filler.ChangeTarget(cell);
                    if (!filler.Check())
                    {
                        continue;
                    }

                    filler.Fill(source);
                }
            }
        }
    }
}
