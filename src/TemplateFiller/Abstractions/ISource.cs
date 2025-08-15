using System.Collections.Generic;

namespace TemplateFiller.Abstractions
{
    /// <summary>
    /// 表示一个键值对形式的数据源，用于数据填充
    /// </summary>
    public interface ISource
    {
        /// <summary>
        /// 获取一项数据
        /// </summary>
        /// <param name="key">数据的键</param>
        /// <returns>数据的值</returns>
        object? this[string key] { get; }
        /// <summary>
        /// 使用特定键，获取数据源的一个节点
        /// </summary>
        /// <param name="key">数据源节点的键</param>
        /// <returns><see cref="ISourceSection"/></returns>
        /// <remarks>
        ///     该方法永远不返回<c>null</c>。如果特定键没有匹配的节点，会返回空的<see cref="ISourceSection"/>
        /// </remarks>
        ISourceSection GetSection(string key);
        /// <summary>
        /// 获取直属的子节点
        /// </summary>
        /// <returns></returns>
        IEnumerable<ISourceSection> GetChildren();
    }
}
