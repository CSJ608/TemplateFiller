namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 表示数据源的一个节点
    /// </summary>
    public interface ISourceSection : ISource
    {
        /// <summary>
        /// 获取该节点在它父级中的键
        /// </summary>
        string Key { get; }
        /// <summary>
        /// 获取该节点在<seealso cref="ISource"/>中的完整路径
        /// </summary>
        string Path { get; }
        /// <summary>
        /// 获取节点的值
        /// </summary>
        object? Value { get; }
    }
}
