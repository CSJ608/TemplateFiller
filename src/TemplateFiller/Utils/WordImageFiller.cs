using FileSignatures;
using NPOI.SS.UserModel;
using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;
using TemplateFiller.Extensions;
using PictureType = NPOI.XWPF.UserModel.PictureType;

namespace TemplateFiller.Utils
{
    /// <summary>
    /// 表示Word图片填充器。
    /// <para>可以识别文档内是否有图片占位符{path}</para>
    /// <para>可以从数据源<seealso cref="ISource"/>中获取source[path]，然后替换占位图片</para>
    /// </summary>
    /// <remarks>
    /// 占位符参见：<seealso cref="PlaceholderConsts.ValuePlaceholder"/>
    /// </remarks>
    public class WordImageFiller(XWPFDocument? document, CancellationToken cancellationToken = default) : ITargetFiller, IDisposable
    {
        private XWPFDocument? Document { get; set; } = document;
        private CancellationToken CancellationToken { get; } = cancellationToken;

        /// <inheritdoc/>
        public bool Check() => CheckHasImage(Document);

        /// <inheritdoc/>
        public void Fill(ISource source)
            => ProcessDocument(Document, source, CancellationToken);

        /// <inheritdoc/>
        public void Dispose()
        {
            Document = null;
            GC.SuppressFinalize(this);
        }

        private static bool CheckHasImage(XWPFDocument? document)
        {
            if(document == null)
            {
                return false;
            }

            return document.AllPictures.Count > 0;
        }

        private static void ProcessDocument(XWPFDocument? document, ISource source, CancellationToken cancellationToken = default)
        {
            if (document == null)
            {
                return;
            }

            ProcessParagraphs(document, document.Paragraphs, source, cancellationToken);

            ProcessTables(document, document.Tables, source, cancellationToken);

            foreach (var header in document.HeaderList)
            {
                ProcessParagraphs(document, header.Paragraphs, source, cancellationToken);
                ProcessTables(document, document.Tables, source, cancellationToken);
            }

            foreach (var footer in document.FooterList)
            {
                ProcessParagraphs(document, footer.Paragraphs, source, cancellationToken);
                ProcessTables(document, document.Tables, source, cancellationToken);
            }
        }

        private static void ProcessTables(XWPFDocument document, IList<XWPFTable> tables, ISource source, CancellationToken cancellationToken)
        {
            foreach (var table in tables)
            {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        ProcessParagraphs(document, cell.Paragraphs, source, cancellationToken);
                    }
                }
            }
        }

        private static void ProcessParagraphs(XWPFDocument document, IList<XWPFParagraph> paragraphs, ISource source, CancellationToken cancellationToken)
        {
            foreach (var paragraph in paragraphs)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var drawCount = paragraph.Runs.SelectMany(t => t.GetCTR().GetDrawingList()).Count();
                var pictureCount = paragraph.Runs.SelectMany(t => t.GetEmbeddedPictures()).Count();
                if (drawCount + pictureCount == 0)
                {
                    continue;
                }

                //foreach (var pic in pictures)
                //{
                //    var picName = pic.GetCTPicture().nvPicPr.cNvPr.name;
                //    if (!picName.IsMatch(PlaceholderConsts.ValuePlaceholder, out var _, out var _))
                //    {
                //        continue;
                //    }

                //    var match = Regex.Match(picName, PlaceholderConsts.ValuePlaceholder); // 与多个占位符匹配时，只处理第一个
                //    var key = match.Groups[1].Value;
                //    var imgData = source[key];
                //    if (imgData is not Stream stream)
                //    {
                //        continue;
                //    }

                //    var width = pic.Width;
                //    var height = pic.Height;

                //    var imgType = GetImgType(source, stream, key);
                //    var pictureType = imgType.ToLower() switch
                //    {
                //        "emf" => PictureType.EMF,
                //        "wmf" => PictureType.WMF,
                //        "pict" => PictureType.PICT,
                //        "jpeg" => PictureType.JPEG,
                //        "png" => PictureType.PNG,
                //        "dib" => PictureType.DIB,
                //        "gif" => PictureType.GIF,
                //        "tiff" => PictureType.TIFF,
                //        "eps" => PictureType.EPS,
                //        "bmp" => PictureType.BMP,
                //        "wpg" => PictureType.WPG,
                //        "svg" => PictureType.SVG,

                //        _ => throw new InvalidOperationException()
                //    };
                //    stream.Seek(0, SeekOrigin.Begin);
                //    var rId = document.AddPictureData(stream, (int)pictureType);
                //    var dpart = document.RelationParts.First(t => t.Relationship.Id == rId).Relationship;
                //    pic.SetPictureReference(dpart);
                //}

                ProcessRuns(document, paragraph.Runs, source);
            }
        }

        private static void ProcessRuns(XWPFDocument document, IList<XWPFRun> runs, ISource source)
        {
            foreach (var run in runs)
            {
                FindDrawingAndFill(source, run);
            }
        }

        private static void FindDrawingAndFill(ISource source, XWPFRun run)
        {
            Stream? imgStream = null;
            string matchedKey = string.Empty;
            int? width = null;
            int? height = null;
            string imgType = string.Empty;

            var ctr = run.GetCTR();

            TryRemoveDrawing(source, ref imgStream, ref matchedKey, ref width, ref height, ref imgType, ctr);            

            if (imgStream == null || width == null || height == null)
            {
                return;
            }

            var pictureType = imgType.ToLower() switch
            {
                "emf" => PictureType.EMF,
                "wmf" => PictureType.WMF,
                "pict" => PictureType.PICT,
                "jpeg" => PictureType.JPEG,
                "png" => PictureType.PNG,
                "dib" => PictureType.DIB,
                "gif" => PictureType.GIF,
                "tiff" => PictureType.TIFF,
                "eps" => PictureType.EPS,
                "bmp" => PictureType.BMP,
                "wpg" => PictureType.WPG,
                "svg" => PictureType.SVG,

                _ => throw new InvalidOperationException()
            };

            imgStream.Seek(0, SeekOrigin.Begin);
            run.AddPicture(imgStream, (int)pictureType, $"{matchedKey}.{imgType}", width.Value, height.Value);
        }

        private static bool TryRemoveDrawing(ISource source, ref Stream? imgStream, ref string matchedKey, ref int? width, ref int? height, ref string imgType, NPOI.OpenXmlFormats.Wordprocessing.CT_R ctr)
        {
            var draws = ctr.GetDrawingList();
            if (draws.Count == 0)
            {
                return false;
            }

            var draw = draws[0]; // 在一个run里，一般只有一个drawing
            bool needRemove = false;
            foreach (var a in draw.anchor)
            {
                var picName = a.docPr.name;
                if (!picName.IsMatch(PlaceholderConsts.ValuePlaceholder, out var _, out var _))
                {
                    continue;
                }

                needRemove = true;
                var match = Regex.Match(picName, PlaceholderConsts.ValuePlaceholder); // 与多个占位符匹配时，只处理第一个
                var key = match.Groups[1].Value;
                var imgData = source[key];
                if (imgData is not Stream s)
                {
                    continue;
                }

                imgStream = s;
                matchedKey = key;
                width = (int)a.extent.cx;
                height = (int)a.extent.cy;

                imgType = GetImgType(source, imgStream, key);
            }

            foreach (var item in draw.inline)
            {
                var picName = item.docPr.name;
                if (!picName.IsMatch(PlaceholderConsts.ValuePlaceholder, out var _, out var _))
                {
                    continue;
                }

                needRemove = true;
                var match = Regex.Match(picName, PlaceholderConsts.ValuePlaceholder); // 与多个占位符匹配时，只处理第一个
                var key = match.Groups[1].Value;
                var imgData = source[key];
                if (imgData is not Stream s)
                {
                    continue;
                }

                imgStream = s;
                matchedKey = key;
                width = (int)item.extent.cx;
                height = (int)item.extent.cy;

                imgType = GetImgType(source, imgStream, key);
            }

            if (needRemove)
            {
                ctr.RemoveDrawing(0);
            }

            return true;
        }

        private static string GetImgType(ISource source, Stream imgStream, string key)
        {
            string imgType;
            if (source is Source s1)
            {
                imgType = s1.GetNestedAttributes(key)
                    ?.OfType<ImgFillAttribute>()
                    .FirstOrDefault()?.PictureType ?? string.Empty;
            }
            else
            {
                var inspector = new FileFormatInspector();
                var format = inspector.DetermineFileFormat(imgStream);
                if (format == null)
                {
                    imgType = "png";
                }
                else
                {
                    imgType = format.Extension;
                }
            }

            return imgType;
        }
    }
}
