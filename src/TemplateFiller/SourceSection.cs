using System;
using System.Collections.Generic;
using TemplateFiller.Abstractions;

namespace TemplateFiller
{
    /// <inheritdoc/>
    public class SourceSection : ISourceSection, IDisposable
    {
        private ISource? Source { get; set; }
        private Source? SectionSource { get; set; }
        private object? ValueCached { get; set; }
        private readonly string _key;
        private readonly string _path;

        /// <summary>
        /// 创建数据源节点
        /// </summary>
        /// <param name="source">数据源</param>
        /// <param name="key">该节点在它父级中的键</param>
        /// <param name="path">该节点在<paramref name="source"/>中的路径</param>
        /// <param name="value">该节点的值</param>
        internal SourceSection(ISource? source, string key, string path, object? value)
        {
            Source = source;
            _key = key;
            _path = path;
            ValueCached = value;

            if (value == null)
            {
                SectionSource = null;
            }
            else
            {
                SectionSource = new Source(value);
            }
        }

        /// <inheritdoc/>
        public object? this[string key] => GetValue(key);

        /// <inheritdoc/>
        public string Key => _key;

        /// <inheritdoc/>
        public string Path => _path;

        /// <inheritdoc/>
        public object? Value => ValueCached;

        /// <inheritdoc/>
        public void Dispose()
        {
            ValueCached = null;
            Source = null;
            if (SectionSource != null)
            {
                SectionSource.Dispose();
                SectionSource = null;
            }

            GC.SuppressFinalize(this);
        }

        private object? GetValue(string key)
        {
            if (SectionSource == null)
            {
                return null;
            }

            return SectionSource[key];
        }

        /// <inheritdoc/>
        public IEnumerable<ISourceSection> GetChildren()
        {
            if (SectionSource == null)
            {
                yield break;
            }

            foreach (var section in SectionSource.GetChildren())
            {
                yield return section;
            }
        }

        /// <inheritdoc/>
        public ISourceSection GetSection(string key)
        {
            if (SectionSource == null)
            {
                return Empty;
            }

            return SectionSource.GetSection(key);
        }

        internal static SourceSection Empty => new(null, string.Empty, string.Empty, null);
    }
}
