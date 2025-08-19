using System;
using System.Collections;
using System.Collections.Generic;
using TemplateFiller.Abstractions;

namespace TemplateFiller
{
    /// <inheritdoc/>
    public class Source(object? source) : ISource, IDisposable
    {
        private readonly Stack<SourceSection> _stack = new();

        /// <inheritdoc/>
        public object? this[string key] => GetNestedValue(key);

        /// <inheritdoc/>
        public IEnumerable<ISourceSection> GetChildren()
        {
            if (source == null)
            {
                yield break;
            }

            if (source is IDictionary dictionary)
            {
                foreach (var childKey in dictionary.Keys)
                {
                    var key = childKey.ToString() ?? string.Empty;
                    if (TryGetNestedValue(source, key, out var value, out _))
                    {
                        yield return CreateSection(key, key, value);
                    }
                }

                yield break;
            }

            var props = source.GetType().GetProperties();
            foreach (var prop in props)
            {
                yield return CreateSection(prop.Name, prop.Name, prop.GetValue(source));
            }
        }

        /// <inheritdoc/>
        public ISourceSection GetSection(string key)
        {
            if (TryGetNestedValue(source, key, out var value, out var sectionKey))
            {
                return CreateSection(sectionKey, key, value);
            }

            return SourceSection.Empty;
        }

        /// <summary>
        /// 尝试获取值
        /// </summary>
        /// <param name="source"></param>
        /// <param name="path"></param>
        /// <param name="value"></param>
        /// <param name="key"></param>
        /// <returns>path有效时，返回true，否则返回false</returns>
        private bool TryGetNestedValue(object? source, string path, out object? value, out string key)
        {
            value = null;
            key = string.Empty;
            if (source == null || string.IsNullOrEmpty(path))
            {
                return false;
            }

            var pathParts = path.Split(':');
            var current = source;

            for (var i = 0; i < pathParts.Length; i++)
            {
                if (current == null)
                {
                    return false;
                }

                var part = pathParts[i];

                // 检查字典
                if (current is IDictionary dictionary)
                {
                    if (dictionary.Contains(part))
                    {
                        current = dictionary[part];
                        continue;
                    }
                    else
                    {
                        return false;
                    }
                }

                // 检查属性
                current = GetPropertyValue(current, part);
            }

            value = current;
            key = pathParts[^1];
            return true;
        }

        private object? GetNestedValue(string path)
        {
            if (TryGetNestedValue(source, path, out var value, out _))
            {
                return value;
            }

            return null;
        }

        private object? GetPropertyValue(object obj, string propertyName)
        {
            var prop = obj.GetType().GetProperty(propertyName);
            return prop?.GetValue(obj);
        }

        private ISourceSection CreateSection(string key, string path, object? value)
        {
            var section = new SourceSection(this, key, path, value);
            _stack.Push(section);
            return section;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            source = null;
            while (_stack.Count > 0)
            {
                var section = _stack.Pop();
                section.Dispose();
            }
        }


    }
}
