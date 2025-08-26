using FileSignatures;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using TemplateFiller.Abstractions;
using TemplateFiller.Extensions;
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

            protected override void FillTemplateImplementation(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            {
                using var source = new Source(dataSource);
                using var workbook = LoadWorkbook(template);
                ProcessWorkbook(workbook, source, cancellationToken);
                workbook.Write(output);
            }

            protected override void FillTemplateImplementation(Stream template, IEnumerable<Bag> bags, CancellationToken cancellationToken = default)
            {
                using var templateWorkbook = LoadWorkbook(template);
                foreach (var bag in bags)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    using var workbook = templateWorkbook.Copy();
                    using var source = new Source(bag.DataSource);
                    ProcessWorkbook(workbook, source, cancellationToken);
                    workbook.Write(bag.Output);
                }
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
