using FileSignatures;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.Util;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
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
            var infos = GetPicturePlaceholders(ctr);

            if (infos.Count == 0)
            {
                return;
            }

            foreach (var info in infos)
            {
                var imgData = source[info.MatchedKey];
                if (imgData is not Stream imgStream)
                {
                    continue;
                }

                BuildPictureTypeAndName(info, imgStream, out var pictureType, out var fileName);

                var parent = run.Parent;
                var picData = BuildPicutreData(document, imgStream, pictureType, parent);

                if (info.IsInline && info.Inline != null)
                {
                    ReplaceInlinePictureData(run, info, fileName, parent, picData);
                }

                if(!info.IsInline && info.Anchor != null)
                {
                    ReplaceAnchorPictureData(run, info, fileName, parent, picData);
                }
            }
        }

        private static void ReplaceInlinePictureData(
            XWPFRun run, 
            PicturePlaceholderInfo info, string fileName, IRunBody parent, XWPFPictureData picData)
        {
            info.Inline!.docPr.descr = fileName;
            info.Inline.graphic.graphicData = new CT_GraphicalObjectData();
            info.Inline.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

            var id = info.Inline.docPr.id;

            // Grab the picture object
            var pic = new NPOI.OpenXmlFormats.Dml.Picture.CT_Picture();

            // Set it up
            var nvPicPr = pic.AddNewNvPicPr();

            var cNvPr = nvPicPr.AddNewCNvPr();
            /* use "0" for the id. See ECM-576, 20.2.2.3 */
            cNvPr.id = 0;
            /* This name is not visible in Word 2010 anywhere */
            cNvPr.name = $"Picture {id}";
            cNvPr.descr = fileName;

            var cNvPicPr = nvPicPr.AddNewCNvPicPr();
            cNvPicPr.AddNewPicLocks().noChangeAspect = true;

            var blipFill = pic.AddNewBlipFill();
            var blip = blipFill.AddNewBlip();
            blip.embed = parent.Part.GetRelationId(picData);
            blipFill.AddNewStretch().AddNewFillRect();

            var spPr = pic.AddNewSpPr();
            var xfrm = spPr.AddNewXfrm();

            var off = xfrm.AddNewOff();
            off.x = (0);
            off.y = (0);

            var ext = xfrm.AddNewExt();
            ext.cx = info.Width;
            ext.cy = info.Height;

            var prstGeom = spPr.AddNewPrstGeom();
            prstGeom.prst = (ST_ShapeType.rect);
            prstGeom.AddNewAvLst();

            using (var ms = RecyclableMemory.GetStream())
            {
                var sw = new StreamWriter(ms);
                pic.Write(sw, "pic:pic");
                sw.Flush();
                ms.Position = 0;
                var sr = new StreamReader(ms);
                var picXml = sr.ReadToEnd();
                info.Inline.graphic.graphicData.AddPicElement(picXml);
            }
            // Finish up
            var xwpfPicture = new XWPFPicture(pic, run);
            var pictures = run.GetEmbeddedPictures();
            pictures.Add(xwpfPicture);
        }

        private static void ReplaceAnchorPictureData(
            XWPFRun run,
            PicturePlaceholderInfo info, string fileName, IRunBody parent, XWPFPictureData picData)
        {
            info.Anchor!.docPr.descr = fileName;
            info.Anchor.graphic.graphicData = new CT_GraphicalObjectData();
            info.Anchor.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";

            var id = info.Anchor.docPr.id;

            // Grab the picture object
            var pic = new NPOI.OpenXmlFormats.Dml.Picture.CT_Picture();

            // Set it up
            var nvPicPr = pic.AddNewNvPicPr();

            var cNvPr = nvPicPr.AddNewCNvPr();
            /* use "0" for the id. See ECM-576, 20.2.2.3 */
            cNvPr.id = 0;
            /* This name is not visible in Word 2010 anywhere */
            cNvPr.name = $"Picture {id}";
            cNvPr.descr = fileName;

            var cNvPicPr = nvPicPr.AddNewCNvPicPr();
            cNvPicPr.AddNewPicLocks().noChangeAspect = true;

            var blipFill = pic.AddNewBlipFill();
            var blip = blipFill.AddNewBlip();
            blip.embed = parent.Part.GetRelationId(picData);
            blipFill.AddNewStretch().AddNewFillRect();

            var spPr = pic.AddNewSpPr();
            var xfrm = spPr.AddNewXfrm();

            var off = xfrm.AddNewOff();
            off.x = (0);
            off.y = (0);

            var ext = xfrm.AddNewExt();
            ext.cx = info.Width;
            ext.cy = info.Height;

            var prstGeom = spPr.AddNewPrstGeom();
            prstGeom.prst = (ST_ShapeType.rect);
            prstGeom.AddNewAvLst();

            using (var ms = RecyclableMemory.GetStream())
            {
                var sw = new StreamWriter(ms);
                pic.Write(sw, "pic:pic");
                sw.Flush();
                ms.Position = 0;
                var sr = new StreamReader(ms);
                var picXml = sr.ReadToEnd();
                info.Anchor.graphic.graphicData.AddPicElement(picXml);
            }
        }

        private static XWPFPictureData BuildPicutreData(XWPFDocument document, Stream imgStream, PictureType pictureType, IRunBody parent)
        {
            String relationId;
            XWPFPictureData picData;
            if (parent.Part is XWPFHeaderFooter headerFooter)
            {
                relationId = headerFooter.AddPictureData(imgStream, (int)pictureType);
                picData = (XWPFPictureData)headerFooter.GetRelationById(relationId);
            }
            else if (parent.Part is XWPFComments comments)
            {
                relationId = comments.AddPictureData(imgStream, (int)pictureType);
                picData = (XWPFPictureData)comments.GetRelationById(relationId);
            }
            else
            {
                relationId = document.AddPictureData(imgStream, (int)pictureType);
                picData = (XWPFPictureData)document.GetRelationById(relationId);
            }

            return picData;
        }

        private static void BuildPictureTypeAndName(PicturePlaceholderInfo info, Stream imgStream, out PictureType pictureType, out string fileName)
        {
            string imgType;
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

            pictureType = imgType.ToLower() switch
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

                _ => throw new InvalidOperationException("Not support")
            };
            imgStream.Seek(0, SeekOrigin.Begin);
            fileName = $"{info.MatchedKey}.{imgType}";
        }

        private class PicturePlaceholderInfo
        {
            public bool IsInline { get; set; }
            public string MatchedKey { get; set; } = string.Empty;
            public int Width { get; set; }
            public int Height { get; set; }
            public CT_Drawing CT_Drawing { get; set; } = new CT_Drawing();
            public CT_Anchor? Anchor { get; set; }
            public CT_Inline? Inline { get; set; }
        }

        private static List<PicturePlaceholderInfo> GetPicturePlaceholders(CT_R ctr)
        {
            var result = new List<PicturePlaceholderInfo>();

            var draws = ctr.GetDrawingList();
            if (draws.Count == 0)
            {
                return result;
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

                    result.Add(new PicturePlaceholderInfo()
                    {
                        IsInline = false,
                        MatchedKey = key,
                        Width = (int)item.extent.cx,
                        Height = (int)item.extent.cy,
                        Anchor = item,
                        CT_Drawing = draw,
                    });
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

                    result.Add(new PicturePlaceholderInfo()
                    {
                        IsInline = true,
                        MatchedKey = key,
                        Width = (int)item.extent.cx,
                        Height = (int)item.extent.cy,
                        Inline = item,
                        CT_Drawing = draw,
                    });
                }
            }

            return result;
        }        
    }
}
