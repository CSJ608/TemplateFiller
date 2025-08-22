using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using TemplateFiller.Abstractions;

using static TemplateFiller.Utils.Default;

namespace TemplateFiller.Extensions
{
    /// <summary>
    /// 填充器批量填充扩展方法
    /// </summary>
    public static class FillerBatchExtensions
    {
        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="bags">作业包</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, string templateFile, IEnumerable<Bag> bags)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            filler.FillTemplate(template, bags);
        }

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="template">模板的流</param>
        /// <param name="fileBags">作业包</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, Stream template, IEnumerable<FileBag> fileBags)
        {
            WithFileBag(fileBags, bags =>
            {
                filler.FillTemplate(template, bags);
            });
        }

        private static IEnumerable<Bag> ConvertToBags(IEnumerable<FileBag> fileBags)
        {
            return fileBags.Select(t =>
                new Bag(
                    OpenWritableFileStream(t.OutputFile),
                    t.DataSource
                )
            );
        }

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="fileBags">作业包</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, string templateFile, IEnumerable<FileBag> fileBags)
        {
            WithFileBag(fileBags, bags =>
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                filler.FillTemplate(template, bags);
            });
        }

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="bags">作业包</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, string templateFile, IEnumerable<Bag> bags,
            CancellationToken cancellationToken = default)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            return filler.FillTemplateAsync(template, bags, cancellationToken);
        }

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="template">模板的流</param>
        /// <param name="fileBags">作业包</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, Stream template, IEnumerable<FileBag> fileBags,
            CancellationToken cancellationToken = default)
        {
            return WithFileBagAsync(fileBags, bags =>
            {
                return filler.FillTemplateAsync(template, bags, cancellationToken);
            });
        }

        /// <summary>
        /// 用同一模板，进行多次填充作业
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="fileBags">作业包</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, string templateFile, IEnumerable<FileBag> fileBags,
            CancellationToken cancellationToken = default)
        {
            return WithFileBagAsync(fileBags, bags =>
            {
                using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
                return filler.FillTemplateAsync(template, bags, cancellationToken);
            });
        }

        private static void WithFileBag(IEnumerable<FileBag> fileBags, Action<IEnumerable<Bag>> action)
        {
            var bags = ConvertToBags(fileBags).ToList();
            try
            {
                action.Invoke(bags);
            }
            finally
            {
                foreach (var bag in bags)
                {
                    bag.Output.Dispose();
                }
            }
        }

        private static Task WithFileBagAsync(IEnumerable<FileBag> fileBags, Func<IEnumerable<Bag>,Task> func)
        {
            var bags = ConvertToBags(fileBags).ToList();
            try
            {
                return func.Invoke(bags);
            }
            finally
            {
                foreach (var bag in bags)
                {
                    bag.Output.Dispose();
                }
            }
        }
    }
}
