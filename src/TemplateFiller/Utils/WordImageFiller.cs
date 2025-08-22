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
    public class WordImageFiller : ITargetFiller, IDisposable
    {
        /// <summary>
        /// 
        /// </summary>
        public WordImageFiller()
        {
            _document = null;
            _paragraph = null;
            FillEmbdata = false;
        }

        private XWPFParagraph? _paragraph { get; set; }

        private XWPFDocument? _document { get; set; }

        public bool FillEmbdata { get; private set; }

        /// <inheritdoc/>
        public bool Check()
        {
            if (_document == null && _paragraph == null)
            {
                return false;
            }

            if(_paragraph != null)
            {
                var draws = _paragraph.Runs.SelectMany(t => t.GetCTR().GetDrawingList()).ToList();
                var picts = _paragraph.Runs.SelectMany(t => t.GetEmbeddedPictures()).ToList();
                if (draws.Count + picts.Count == 0) { 
                    return false;
                }

                foreach (var anchor in draws.SelectMany(t => t.anchor))
                {
                    anchor.graphic.AddNewGraphicData
                }

                var graphics = draws.SelectMany(t => t.anchor).Select(t => t.graphic).ToList();
                foreach (var item in graphics)
                {
                    var drawName = item.AddNewGraphicData();
                    drawName.
                }

                foreach(var pic in picts)
                {
                    var picName = pic.GetCTPicture().nvPicPr.cNvPr.name;
                }
            }

            return false;
        }

        /// <inheritdoc/>
        public void Fill(ISource source)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 将目标切换为段落内的嵌入式图片
        /// </summary>
        /// <param name="paragraph"></param>
        public void ChangeTarget(XWPFParagraph paragraph)
        {
            _paragraph = paragraph;
            _document = null;
            FillEmbdata = true;
        }

        /// <summary>
        /// 将目标切换为文档内的非嵌入式图片
        /// </summary>
        /// <param name="document"></param>
        public void ChangeTarget(XWPFDocument document)
        {
            _document = document;
            _paragraph = null;
            FillEmbdata = false;
        }

        public void Dispose()
        {
            
        }
    }
}
