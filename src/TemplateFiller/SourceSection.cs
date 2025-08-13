using System;
using System.Collections.Generic;
using TemplateFiller.Abstractions;

namespace TemplateFiller
{
    public class SourceSection : ISourceSection, IDisposable
    {
        private ISource? _source { get; set; }
        private Source? _sectionSource { get; set; }
        private object? _value { get; set; }
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
            _source = source;
            _key = key;
            _path = path;
            _value = value;

            if (value == null)
            {
                _sectionSource = null;
            }
            else
            {
                _sectionSource = new Source(value);
            }
        }

        public object? this[string key] => GetValue(key);

        public string Key => _key;

        public string Path => _path;

        public object? Value => _value;

        public void Dispose()
        {
            _value = null;
            _source = null;
            if (_sectionSource != null)
            {
                _sectionSource.Dispose();
                _sectionSource = null;
            }
        }

        private object? GetValue(string key)
        {
            if (_sectionSource == null)
            {
                return null;
            }

            return _sectionSource[key];
        }

        public IEnumerable<ISourceSection> GetChildren()
        {
            if (_sectionSource == null)
            {
                yield break;
            }

            foreach (var section in _sectionSource.GetChildren())
            {
                yield return section;
            }
        }

        public ISourceSection GetSection(string key)
        {
            if (_sectionSource == null)
            {
                return Empty;
            }

            return _sectionSource.GetSection(key);
        }

        public static SourceSection Empty => new SourceSection(null, string.Empty, string.Empty, null);
    }
}
