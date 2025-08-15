using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;
using TemplateFiller.Abstractions;
using TemplateFiller.Utils;

namespace TemplateFiller
{
    public class NpoiWordTemplateFiller
    {
        public static bool CreateDirectoryIfNotExists { get; set; } = true;
        public void FillTemplate(string templatePath, string outputPath, object dataSource)
        {
            using (var source = new Source(dataSource))
            {
                using var doc = LoadDocument(templatePath);
                ProcessDocument(doc, source);
                using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    doc.Write(stream);
                }
            }
        }

        private XWPFDocument LoadDocument(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                return new XWPFDocument(stream);
            }
        }

        private void ProcessDocument(XWPFDocument doc, ISource source)
        {
            using var filler = new WordValueFiller(null);

            // 处理文档
            ProcessParagraph(source, filler, doc.Paragraphs);

            // 处理表格
            ProcessTables(source, filler, doc.Tables);

            // 处理页眉页脚
            foreach (var header in doc.HeaderList)
            {
                ProcessParagraph(source, filler, header.Paragraphs);
            }
            foreach (var footer in doc.FooterList)
            {
                ProcessParagraph(source, filler, footer.Paragraphs);
            }
        }

        private static void ProcessTables(ISource source, WordValueFiller filler, IList<XWPFTable> tables)
        {
            using var arrayfiller = new WordArrayFiller(null);
            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        ProcessParagraph(source, filler, cell.Paragraphs);
                    }
                }

                arrayfiller.ChangeTarget(table);
                if (!arrayfiller.Check())
                {
                    continue;
                }
                arrayfiller.Fill(source);
            }
        }

        private static void ProcessParagraph(ISource source, WordValueFiller filler, IList<XWPFParagraph> paragraphs)
        {
            // 处理段落中的文本
            foreach (var paragraph in paragraphs)
            {
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
