using System;
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
        /// 填充模板的实现类
        /// </summary>
        /// <param name="template"></param>
        /// <param name="output"></param>
        /// <param name="dataSource"></param>
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

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static NotSupportedException ThrowNotSupportException(TemplateType templateType, [CallerMemberName] string? memberName = null)
        => new($"{memberName} is not supported, since the templateFiller's {nameof(TemplateType)} is {templateType}");

        private protected Filler(TemplateType templateType)
        {
            TemplateType = templateType;
        }
    }
}
