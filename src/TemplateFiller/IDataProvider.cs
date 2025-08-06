using System;
using System.Collections.Generic;

namespace TemplateFiller
{
    public interface IDataProvider
    {
        object? this[string key] { get; }
        IEnumerable<IDataProvider> GetChildren(string key);
        bool TryGetValue(string key, out object? value);
    }
}
