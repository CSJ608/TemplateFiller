using System.IO;
using System.Threading;
using System.Threading.Tasks;

using static TemplateFiller.Utils.Default;

namespace TemplateFiller.Extensions
{
    /// <summary>
    /// 填充器扩展方法
    /// </summary>
    public static class FillerExtensions
    {
        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, string templateFile, Stream output, object dataSource)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            filler.FillTemplate(template, output, dataSource);
        }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="template">模板的流</param>
        /// <param name="outputFile">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, Stream template, string outputFile, object dataSource)
        {
            using var output = OpenWritableFileStream(outputFile);
            filler.FillTemplate(template, output, dataSource);
        }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="outputFile">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <returns></returns>
        public static void FillTemplate(this Filler filler, string templateFile, string outputFile, object dataSource)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            using var output = OpenWritableFileStream(outputFile);
            filler.FillTemplate(template, output, dataSource);
        }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="output">输出流</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, string templateFile, Stream output, object dataSource, CancellationToken cancellationToken = default)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            return filler.FillTemplateAsync(template, output, dataSource, cancellationToken);
        }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="template">模板的流</param>
        /// <param name="outputFile">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, Stream template, string outputFile, object dataSource, CancellationToken cancellationToken = default)
        {
            using var output = OpenWritableFileStream(outputFile);
            return filler.FillTemplateAsync(template, output, dataSource, cancellationToken);
        }

        /// <summary>
        /// 填充模板
        /// </summary>
        /// <param name="filler"></param>
        /// <param name="templateFile">模板文件</param>
        /// <param name="outputFile">输出文件</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static Task FillTemplateAsync(this Filler filler, string templateFile, string outputFile, object dataSource, CancellationToken cancellationToken = default)
        {
            using var template = new FileStream(templateFile, FileMode.Open, FileAccess.Read);
            using var output = OpenWritableFileStream(outputFile);
            return filler.FillTemplateAsync(template, output, dataSource, cancellationToken);
        }
    }
}
