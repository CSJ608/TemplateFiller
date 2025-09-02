using FileSignatures;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.Util;
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

            return document.AllPictures.Count + document.AllPackagePictures.Count > 0;
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
                ProcessTables(document, header.Tables, source, cancellationToken);
            }

            foreach (var footer in document.FooterList)
            {
                ProcessParagraphs(document, footer.Paragraphs, source, cancellationToken);
                ProcessTables(document, footer.Tables, source, cancellationToken);
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

                ProcessRuns(document, paragraph.Runs, source);
            }
        }

        private static void ProcessRuns(XWPFDocument document, IList<XWPFRun> runs, ISource source)
        {
            foreach (var run in runs)
            {
                FindDrawingAndFill(document, source, run);
            }
        }

        private static void FindDrawingAndFill(XWPFDocument document, ISource source, XWPFRun run)
        {
            var ctr = run.GetCTR();
            var info = GetPicturePlaceholderInfo(ctr);

            if (info == null)
            {
                return;
            }

            for (int i = 0; i < ctr.GetDrawingList().Count; i++)
            {
                ctr.RemoveDrawing(0);
            }

            var newDrawing = ctr.AddNewDrawing();
            var newInline = newDrawing.AddNewInline();
            if (info.IsInline && info.Inline != null)
            {
                newInline.extent = new CT_PositiveSize2D()
                {
                    cx = info.Inline.extent.cx,
                    cy = info.Inline.extent.cy,
                };

                newInline.effectExtent = new CT_EffectExtent()
                {
                    l = info.Inline.effectExtent.l,
                    t = info.Inline.effectExtent.t,
                    r = info.Inline.effectExtent.r,
                    b = info.Inline.effectExtent.b
                };

                var ext = newInline.AddNewExtent();
                ext.cx = info.Width;
                ext.cy = info.Height;
                var pr = newInline.AddNewDocPr();
                pr.add
            }

            //var imgData = source[info.MatchedKey];
            //if (imgData is not Stream s)
            //{
            //    return; //系统
            //}

            //var imgStream = s;
            //string imgType = string.Empty;
            //var inspector = new FileFormatInspector();
            //var format = inspector.DetermineFileFormat(imgStream);
            //if (format == null)
            //{
            //    imgType = "png";
            //}
            //else
            //{
            //    imgType = format.Extension;
            //}

            //var pictureType = imgType.ToLower() switch
            //{
            //    "emf" => PictureType.EMF,
            //    "wmf" => PictureType.WMF,
            //    "pict" => PictureType.PICT,
            //    "jpeg" => PictureType.JPEG,
            //    "png" => PictureType.PNG,
            //    "dib" => PictureType.DIB,
            //    "gif" => PictureType.GIF,
            //    "tiff" => PictureType.TIFF,
            //    "eps" => PictureType.EPS,
            //    "bmp" => PictureType.BMP,
            //    "wpg" => PictureType.WPG,
            //    "svg" => PictureType.SVG,

            //    _ => throw new InvalidOperationException("Not support")
            //};

            //imgStream.Seek(0, SeekOrigin.Begin);
            //run.AddPicture(imgStream, (int)pictureType, $"{matchedKey}.{imgType}", width.Value, height.Value);
        }

        private class PicturePlaceholderInfo
        {
            public bool IsInline { get; set; }
            public string MatchedKey { get; set; } = string.Empty;
            public int Width { get; set; }
            public int Height { get; set; }
            public CT_Anchor? Anchor { get; set; }
            public CT_Inline? Inline { get; set; }
        }

        private static PicturePlaceholderInfo? GetPicturePlaceholderInfo(CT_R ctr)
        {
            var draws = ctr.GetDrawingList();
            if (draws.Count == 0)
            {
                return null;
            }

            foreach (var draw in draws)
            {
                foreach (var item in draw.anchor)
                {
                    var picName = item.docPr.name;
                    if (!picName.IsMatch(PlaceholderConsts.ValuePlaceholder, out var _, out var _))
                    {
                        continue;
                    }

                    var match = Regex.Match(picName, PlaceholderConsts.ValuePlaceholder); // 与多个占位符匹配时，只处理第一个
                    var key = match.Groups[1].Value;

                    return new PicturePlaceholderInfo()
                    {
                        IsInline = false,
                        MatchedKey = key,
                        Width = (int)item.extent.cx,
                        Height = (int)item.extent.cy,
                        Anchor = item,
                    };
                }

                foreach (var item in draw.inline)
                {
                    var picName = item.docPr.name;
                    if (!picName.IsMatch(PlaceholderConsts.ValuePlaceholder, out var _, out var _))
                    {
                        continue;
                    }

                    var match = Regex.Match(picName, PlaceholderConsts.ValuePlaceholder); // 与多个占位符匹配时，只处理第一个
                    var key = match.Groups[1].Value;

                    return new PicturePlaceholderInfo()
                    {
                        IsInline = true,
                        MatchedKey = key,
                        Width = (int)item.extent.cx,
                        Height = (int)item.extent.cy,
                        Inline = item,
                    };
                }
            }

            return null;
        }        
    }
}
