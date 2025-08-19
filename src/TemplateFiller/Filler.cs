using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using TemplateFiller.Abstractions;

namespace TemplateFiller
{
    public abstract partial class Filler
    {
        public TemplateType TemplateType { get; }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual void FillTemplate(Stream template, Stream output, object dataSource)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板文件</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual void FillTemplate(string templateFile, Stream output, object dataSource)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual void FillTemplate(Stream template, string outputFile, object dataSource)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板文件</param>
        /// <param name="output">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual void FillTemplate(string templateFile, string outputFile, object dataSource)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public virtual Task FillTemplateAsync(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板文件</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual Task FillTemplateAsync(string templateFile, Stream output, object dataSource, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual Task FillTemplateAsync(Stream template, string outputFile, object dataSource, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板文件</param>
        /// <param name="output">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public virtual Task FillTemplateAsync(string templateFile, string outputFile, object dataSource, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static NotSupportedException ThrowNotSupportException(TemplateType templateType, [CallerMemberName] string? memberName = null)
        => new($"{memberName} is not supported, since the templateFiller's {nameof(TemplateType)} is {templateType}");

        private protected Filler(TemplateType templateType)
        {
            TemplateType = templateType;
        }
    }
}
