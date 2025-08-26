using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 标识在填充图片占位符时的特性
    /// </summary>
    /// <remarks>
    /// 图片格式为<paramref name="pictureType"/>
    /// </remarks>
    /// <param name="pictureType">图片格式。例如：png、jpeg</param>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ImgFillAttribute(string pictureType = "png") : Attribute
    {
        /// <summary>
        /// 图片格式
        /// </summary>
        public string PictureType { get; } = pictureType;
    }
}
