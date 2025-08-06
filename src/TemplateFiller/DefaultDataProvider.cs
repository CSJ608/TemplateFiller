using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller
{
    public class DefaultDataProvider : IDataProvider, IDisposable
    {
        private readonly object _data;
        private readonly string? _prefix;

        public DefaultDataProvider(object data, string? prefix = null)
        {
            _data = data;
            _prefix = prefix;
        }

        public object? this[string key] => GetValue(key);

        public IEnumerable<IDataProvider> GetChildren(string key)
        {
            var value = GetNestedValue(key);
            if (value is IEnumerable enumerable && !(value is string))
            {
                foreach (var item in enumerable)
                {
                    yield return new DefaultDataProvider(item);
                }
            }
        }

        public bool TryGetValue(string key, out object? value)
        {
            value = GetValue(key);
            return value != null;
        }

        private object? GetValue(string key)
        {
            if (!string.IsNullOrEmpty(_prefix))
            {
                key = $"{_prefix}:{key}";
            }
            return GetNestedValue(key);
        }

        private object? GetNestedValue(string propertyPath)
        {
            if (_data == null || string.IsNullOrEmpty(propertyPath))
                return null;

            var pathParts = propertyPath.Split(':');
            var current = _data;

            for (int i = 0; i < pathParts.Length; i++)
            {
                if (current == null) return null;

                var part = pathParts[i];

                // 优先检查字典（即使当前是IEnumerable）
                if (current is IDictionary dictionary && dictionary.Contains(part))
                {
                    current = dictionary[part];
                    continue;
                }

                // 如果不是字典或是字典但不包含该键，尝试作为普通属性
                var value = GetPropertyValue(current, part);

                // 如果当前不是集合或者不是路径的最后一部分，继续处理
                if (!(value is IEnumerable enumerable) || value is string || i < pathParts.Length - 1)
                {
                    current = value;
                    continue;
                }

                // 只有当明确是集合路径时才返回整个集合
                return enumerable;
            }

            return current;
        }

        private object? GetPropertyValue(object obj, string propertyName)
        {
            if (obj is IDictionary dictionary)
            {
                return dictionary.Contains(propertyName) ? dictionary[propertyName] : null;
            }

            var prop = obj.GetType().GetProperty(propertyName);
            return prop?.GetValue(obj);
        }

        public void Dispose()
        {
            
        }
    }
}
