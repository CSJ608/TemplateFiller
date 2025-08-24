using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TemplateFiller.Abstractions;
using TemplateFiller.Consts;

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
    public class WordImageFiller(XWPFDocument? document) : ITargetFiller, IDisposable
    {
        private XWPFDocument? Document { get; set; } = document;

        /// <inheritdoc/>
        public bool Check() => true;

        /// <inheritdoc/>
        public void Fill(ISource source)
        {
            throw new NotImplementedException();
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            Document = null;
            GC.SuppressFinalize(this);
        }

        private static void CheckHasImagePlaceholder(XWPFDocument? document)
        {
            if (document == null)
            {
                return;
            }

            var draws = paragraph.Runs.SelectMany(t => t.GetCTR().GetDrawingList()).ToList();
            var picts = paragraph.Runs.SelectMany(t => t.GetEmbeddedPictures()).ToList();
            if (draws.Count + picts.Count == 0)
            {
                return;
            }

            //foreach (var anchor in draws.SelectMany(t => t.anchor))
            //{
            //    anchor.graphic.AddNewGraphicData
            //    }

            //var graphics = draws.SelectMany(t => t.anchor).Select(t => t.graphic).ToList();
            //foreach (var item in graphics)
            //{
            //    var drawName = item.AddNewGraphicData();
            //    drawName.
            //    }

            foreach (var pic in picts)
            {
                var picName = pic.GetCTPicture().nvPicPr.cNvPr.name;
                var test = document.GetRelationId(pic.GetPictureData());
            }

            return false;
        }
    }
}
