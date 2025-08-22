using System;

namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 进行一次填充作业用到的包
    /// </summary>
    public class FileBag(string outputFile, object dataSource)
    {
        /// <summary>
        /// 输出文件
        /// </summary>
        public string OutputFile { get; } = outputFile ?? throw new ArgumentNullException(nameof(outputFile));
        /// <summary>
        /// 数据源
        /// </summary>
        public object DataSource { get; } = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
    }
}
