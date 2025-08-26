using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using TemplateFiller.Abstractions;

namespace TemplateFiller
{
    /// <summary>
    /// 填充器
    /// </summary>
    public abstract partial class Filler
    {
        /// <summary>
        /// 模板类型
        /// </summary>
        public TemplateType TemplateType { get; }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        /// <remarks>
        ///     读取模版后，会关闭流<paramref name="template"/>。
        /// </remarks>
        public void FillTemplate(Stream template, Stream output, object dataSource)
            => FillTemplateImplementation(template, output, dataSource);

        /// <summary>
        /// 填充模板的实现
        /// </summary>
        /// <param name="template">模板流</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        protected virtual void FillTemplateImplementation(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public Task FillTemplateAsync(Stream template, Stream output, object dataSource, CancellationToken cancellationToken = default)
        {
            FillTemplateImplementation(template, output, dataSource, cancellationToken);
            return Task.CompletedTask;
        }

        /// <summary>
        /// 批量填充模板的实现
        /// </summary>
        /// <param name="template">模板流</param>
        /// <param name="bags">作业包</param>
        /// <param name="cancellationToken"></param>
        protected virtual void FillTemplateImplementation(Stream template, IEnumerable<Bag> bags, CancellationToken cancellationToken = default)
            => throw ThrowNotSupportException(TemplateType);

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="bags">作业包</param>
        /// <returns></returns>
        /// <remarks>
        ///     读取模版后，会关闭流<paramref name="template"/>。
        /// </remarks>
        public void FillTemplate(Stream template, IEnumerable<Bag> bags)
            => FillTemplateImplementation(template, bags);

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="template">模板的流</param>
        /// <param name="bags">作业包</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        /// <remarks>
        ///     读取模版后，会关闭流<paramref name="template"/>。
        /// </remarks>
        public Task FillTemplateAsync(Stream template, IEnumerable<Bag> bags, CancellationToken cancellationToken = default)
        {
            FillTemplateImplementation(template, bags, cancellationToken);
            return Task.CompletedTask;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static NotSupportedException ThrowNotSupportException(TemplateType templateType, [CallerMemberName] string? memberName = null)
        => new($"{memberName} is not supported, since the templateFiller's {nameof(TemplateType)} is {templateType}");

        private protected Filler(TemplateType templateType)
        {
            TemplateType = templateType;
        }
    }
}
