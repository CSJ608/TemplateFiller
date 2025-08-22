using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller.Abstractions
{

    /// <summary>
    /// 标识在填充图片占位符时的特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImgFillAttribute : Attribute
    {
        /// <summary>
        /// 缩放图片使其和占位图片一样大
        /// </summary>
        public bool ScaleToFit { get; set; }

        /// <summary>
        /// 标识在填充图片占位符时的特性
        /// </summary>
        /// <param name="scaleToFit">是否缩放到和占位图片一样大</param>
        public ImgFillAttribute(bool scaleToFit = false)
        {
            ScaleToFit = scaleToFit;
        }
    }
}
