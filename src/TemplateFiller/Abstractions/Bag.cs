using System;
using System.IO;

namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 进行一次填充作业用到的包
    /// </summary>
    public class Bag(Stream output, object dataSource)
    {
        /// <summary>
        /// 输出流
        /// </summary>
        public Stream Output { get; } = output ?? throw new ArgumentNullException(nameof(output));
        /// <summary>
        /// 数据源
        /// </summary>
        public object DataSource { get; } = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
    }
}
