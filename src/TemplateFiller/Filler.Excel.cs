using FileSignatures;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using TemplateFiller.Abstractions;
using TemplateFiller.Utils;

namespace TemplateFiller
{
    partial class Filler
    {
        private sealed class ExcelFiller : Filler
        {
            internal ExcelFiller() : base(TemplateType.Excel)
            {
            }

            public override void FillTemplate(Stream template, Stream output, object dataSource)
            {
                using var ms = new MemoryStream();
                template.CopyTo(ms);
                FillTemplateImplementation(ms, output, dataSource);
            }
            
            public override void FillTemplate(string templateFile, Stream output, object dataSource)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                FillTemplateImplementation(template, output, dataSource);
            }
            public override void FillTemplate(Stream template, string outputFile, object dataSource)
            {
                using var ms = new MemoryStream();
                template.CopyTo(ms);
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(ms, output, dataSource);
            }
            public override void FillTemplate(string templateFile, string outputFile, object dataSource)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(template, output, dataSource);
            }
            public override Task FillTemplateAsync(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            {
                using var ms = new MemoryStream();
                template.CopyTo(ms);
                FillTemplateImplementation(ms, output, dataSource, cancellationToken);
                return Task.CompletedTask;
            }
            public override Task FillTemplateAsync(string templateFile, Stream output, object dataSource, CancellationToken cancellationToken = default)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                FillTemplateImplementation(template, output, dataSource, cancellationToken);
                return Task.CompletedTask;
            }
            public override Task FillTemplateAsync(Stream template, string outputFile, object dataSource, CancellationToken cancellationToken = default)
            {
                using var ms = new MemoryStream();
                template.CopyTo(ms);
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(ms, output, dataSource, cancellationToken);
                return Task.CompletedTask;
            }
            public override Task FillTemplateAsync(string templateFile, string outputFile, object dataSource, CancellationToken cancellationToken = default)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(template, output, dataSource, cancellationToken);
                return Task.CompletedTask;
            }

            private static Stream OpenOutputStream(string outputFile)
            {
                var directory = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                return new FileStream(outputFile, FileMode.Create, FileAccess.Write);
            }

            private static void FillTemplateImplementation(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            {
                using var source = new Source(dataSource);
                using var workbook = LoadWorkbook(template);
                ProcessWorkbook(workbook, source, cancellationToken);
                workbook.Write(output);
            }

            private static IWorkbook LoadWorkbook(Stream stream)
            {
                if (stream == null) throw new ArgumentNullException(nameof(stream));
                if (!stream.CanRead) throw new ArgumentException("Stream is not readable", nameof(stream));
                if (!stream.CanSeek) throw new ArgumentException("Stream must be seekable", nameof(stream));

                var inspector = new FileFormatInspector();
                var format = inspector.DetermineFileFormat(stream);
                if (format == null || format.Extension != "xls" && format.Extension != "xlsx")
                {
                    throw new NotSupportedException($"File format '{format?.Extension ?? string.Empty}' is not supported. Supported formats: .xls, .xlsx");
                }

                return format.Extension.ToLower() switch
                {
                    "xls" => new HSSFWorkbook(stream),
                    "xlsx" => new XSSFWorkbook(stream),
                    _ => throw new NotSupportedException($"File format '{format?.Extension ?? string.Empty}' is not supported. Supported formats: .xls, .xlsx"),
                };
            }
            private static void ProcessWorkbook(IWorkbook workbook, ISource source, CancellationToken cancellationToken = default)
            {
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    ProcessSheet(sheet, source, cancellationToken);
                }
            }

            private static void ProcessSheet(ISheet sheet, ISource source, CancellationToken cancellationToken = default)
            {
                // 先处理可能包含数组占位符的单元格
                ProcessArrayPlaceholders(sheet, source, cancellationToken);

                // 然后处理普通占位符
                ProcessRegularPlaceholders(sheet, source, cancellationToken);
            }

            private static void ProcessArrayPlaceholders(ISheet sheet, ISource source, CancellationToken cancellationToken = default)
            {
                var filler = new ExcelArrayFiller(null);
                for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

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

            private static void ProcessRegularPlaceholders(ISheet sheet, ISource source, CancellationToken cancellationToken = default)
            {
                var filler = new ExcelValueFiller(null);
                for (int rowNum = sheet.FirstRowNum; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

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
}
