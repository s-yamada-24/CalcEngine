using System;
using System.Collections.Generic;

namespace CalcEngine
{
    public class FormulaEvaluator
    {
        private readonly VirtualTable _table;
        public FunctionRegistry FunctionRegistry { get; }
        
        // Cache for parsed ASTs
        private readonly Dictionary<string, AstNode> _astCache = new Dictionary<string, AstNode>();

        public FormulaEvaluator(VirtualTable table)
        {
            _table = table;
            FunctionRegistry = new FunctionRegistry();
            StandardFunctions.Register(FunctionRegistry);
        }

        public object Evaluate(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return null;

            string expression = formula;
            if (formula.StartsWith("="))
            {
                expression = formula.Substring(1);
            }

            // Check cache first
            if (!_astCache.TryGetValue(expression, out var ast))
            {
                var lexer = new Lexer(expression);
                var tokens = lexer.Tokenize();
                var parser = new Parser(tokens);
                ast = parser.Parse();
                _astCache[expression] = ast;
            }

            var evaluator = new Evaluator(_table, FunctionRegistry);
            return evaluator.Evaluate(ast);
        }

        public void ClearCache()
        {
            _astCache.Clear();
        }
    }
}
