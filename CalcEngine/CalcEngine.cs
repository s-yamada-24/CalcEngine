using System;

namespace CalcEngine
{
    public class CalcEngine
    {
        private readonly VirtualTable _table;
        private readonly FormulaEvaluator _evaluator;

        public CalcEngine()
        {
            _table = new VirtualTable();
            _evaluator = new FormulaEvaluator(_table);
        }

        /// <summary>
        /// Sets a value in the virtual table.
        /// </summary>
        /// <param name="address">The cell address (e.g., "A1").</param>
        /// <param name="value">The value to set.</param>
        public void SetValue(string address, object? value)
        {
            _table.SetValue(address, value);
        }

        /// <summary>
        /// Gets a value from the virtual table.
        /// </summary>
        /// <param name="address">The cell address (e.g., "A1").</param>
        /// <returns>The value in the cell.</returns>
        public object? GetValue(string address)
        {
            return _table.GetValue(address);
        }

        /// <summary>
        /// Evaluates a formula string.
        /// </summary>
        /// <param name="formula">The formula to evaluate (e.g., "=A1+B1").</param>
        /// <returns>The result of the evaluation.</returns>
        public object Evaluate(string formula)
        {
            return _evaluator.Evaluate(formula);
        }

        /// <summary>
        /// Registers a custom function.
        /// </summary>
        /// <param name="name">The function name.</param>
        /// <param name="function">The function delegate.</param>
        public void RegisterFunction(string name, FunctionDelegate function)
        {
            _evaluator.FunctionRegistry.Register(name, function);
        }

        /// <summary>
        /// Clears all values from the virtual table.
        /// </summary>
        public void Clear()
        {
            _table.Clear();
        }
    }
}
