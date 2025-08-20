using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using TemplateFiller.Abstractions;
using TemplateFiller.Utils;

namespace TemplateFiller
{
    partial class Filler
    {
        private sealed class WordFiller : Filler
        {
            internal WordFiller() : base(TemplateType.Word)
            {

            }

            public override void FillTemplate(Stream template, Stream output, object dataSource)
                => FillTemplateImplementation(template, output, dataSource);
            public override void FillTemplate(string templateFile, Stream output, object dataSource)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                FillTemplateImplementation(template, output, dataSource);
            }
            public override void FillTemplate(Stream template, string outputFile, object dataSource)
            {
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(template, output, dataSource);
            }
            public override void FillTemplate(string templateFile, string outputFile, object dataSource)
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(template, output, dataSource);
            }
            public override Task FillTemplateAsync(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            {
                FillTemplateImplementation(template, output, dataSource, cancellationToken);
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
                using var output = OpenOutputStream(outputFile);
                FillTemplateImplementation(template, output, dataSource, cancellationToken);
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
                using var doc = new XWPFDocument(template);
                ProcessDocument(doc, source, cancellationToken);
                doc.Write(output);
            }

            private static void ProcessDocument(XWPFDocument doc, ISource source, CancellationToken cancellationToken = default)
            {
                using var filler = new WordValueFiller(null);

                // 处理文档
                ProcessParagraph(source, filler, doc.Paragraphs, cancellationToken);

                // 处理表格
                ProcessTables(source, filler, doc.Tables, true, cancellationToken);

                // 处理页眉页脚
                foreach (var header in doc.HeaderList)
                {
                    ProcessParagraph(source, filler, header.Paragraphs, cancellationToken);
                    ProcessTables(source, filler, header.Tables, false, cancellationToken);
                }
                foreach (var footer in doc.FooterList)
                {
                    ProcessParagraph(source, filler, footer.Paragraphs, cancellationToken);
                    ProcessTables(source, filler, footer.Tables, false, cancellationToken);
                }
            }

            private static void ProcessTables(ISource source, WordValueFiller filler, IList<XWPFTable> tables, bool fillArray, CancellationToken cancellationToken = default)
            {
                using var arrayfiller = new WordArrayFiller(null);
                foreach (var table in tables)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.GetTableCells())
                        {
                            ProcessParagraph(source, filler, cell.Paragraphs, cancellationToken);
                        }
                    }

                    if (!fillArray)
                    {
                        continue;
                    }

                    arrayfiller.ChangeTarget(table);
                    if (!arrayfiller.Check())
                    {
                        continue;
                    }
                    arrayfiller.Fill(source);
                }
            }

            private static void ProcessParagraph(ISource source, WordValueFiller filler, IList<XWPFParagraph> paragraphs, CancellationToken cancellationToken = default)
            {
                // 处理段落中的文本
                foreach (var paragraph in paragraphs)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    filler.ChangeTarget(paragraph);
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
