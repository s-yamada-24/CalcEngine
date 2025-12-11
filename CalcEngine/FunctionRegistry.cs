using System;
using System.Collections.Generic;

namespace CalcEngine
{
    public delegate object FunctionDelegate(object[] args);

    public class FunctionRegistry
    {
        private readonly Dictionary<string, FunctionDelegate> _functions = new(StringComparer.OrdinalIgnoreCase);

        public void Register(string name, FunctionDelegate function)
        {
            _functions[name] = function;
        }

        public object Call(string name, object[] args)
        {
            if (_functions.TryGetValue(name, out var func))
            {
                return func(args);
            }
            throw new KeyNotFoundException($"Function '{name}' not found");
        }
    }
}
