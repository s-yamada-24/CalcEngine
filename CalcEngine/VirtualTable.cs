using System;
using System.Collections.Generic;

namespace CalcEngine
{
    public class VirtualTable
    {
        private readonly Dictionary<string, object?> _cells = new(StringComparer.OrdinalIgnoreCase);

        public void SetValue(string address, object? value)
        {
            _cells[address] = value;
        }

        public object? GetValue(string address)
        {
            return _cells.TryGetValue(address, out var value) ? value : null;
        }

        public void Clear()
        {
            _cells.Clear();
        }
    }
}
