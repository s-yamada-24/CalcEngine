using System;
using System.Collections.Generic;
using System.Linq;

namespace CalcEngine
{
    public class Evaluator
    {
        private readonly VirtualTable _table;
        private readonly FunctionRegistry _registry;

        public Evaluator(VirtualTable table, FunctionRegistry registry)
        {
            _table = table;
            _registry = registry;
        }

        public object Evaluate(AstNode node)
        {
            switch (node)
            {
                case NumberNode n: return n.Value;
                case StringNode s: return s.Value;
                case CellReferenceNode c: return GetCellValue(c.Address);
                case RangeNode r: return GetRangeValues(r.StartAddress, r.EndAddress);
                case BinaryOpNode b: return EvaluateBinary(b);
                case UnaryOpNode u: return EvaluateUnary(u);
                case FunctionCallNode f: return EvaluateFunction(f);
                default: return CalcError.Value;
            }
        }

        private object GetCellValue(string address)
        {
            var val = _table.GetValue(address);
            if (val == null) return CalcError.Ref; // Return #REF! for missing cells
            if (val is CalcError) return val;
            return val;
        }

        private object GetRangeValues(string start, string end)
        {
            var cells = new List<object>();
            try
            {
                var (startCol, startRow) = ParseAddress(start);
                var (endCol, endRow) = ParseAddress(end);

                int minCol = Math.Min(startCol, endCol);
                int maxCol = Math.Max(startCol, endCol);
                int minRow = Math.Min(startRow, endRow);
                int maxRow = Math.Max(startRow, endRow);

                for (int r = minRow; r <= maxRow; r++)
                {
                    for (int c = minCol; c <= maxCol; c++)
                    {
                        var addr = GetAddress(c, r);
                        // For ranges, we usually want values. 
                        // If a cell is missing in a range, standard Excel behavior varies by function.
                        // But strictly following "missing = #REF!" might be annoying for ranges.
                        // However, to be consistent with GetCellValue:
                        var val = _table.GetValue(addr);
                        // In ranges, empty cells are often treated as 0 or ignored by functions like SUM.
                        // But if we strictly return #REF!, SUM(#REF!) becomes #REF!.
                        // Let's return the value (or null) and let functions handle it.
                        // Wait, user said "If non-existent address is specified... output reference error".
                        // So we should probably include #REF! in the list if it's missing.
                        if (val == null) cells.Add(CalcError.Ref);
                        else cells.Add(val);
                    }
                }
            }
            catch
            {
                return CalcError.Ref;
            }
            return cells;
        }

        private (int col, int row) ParseAddress(string address)
        {
            int i = 0;
            while (i < address.Length && char.IsLetter(address[i])) i++;
            string colStr = address.Substring(0, i);
            string rowStr = address.Substring(i);
            
            int col = 0;
            foreach (char c in colStr.ToUpper())
            {
                col = col * 26 + (c - 'A' + 1);
            }
            
            if (!int.TryParse(rowStr, out int row))
                throw new ArgumentException();
                
            return (col, row);
        }

        private string GetAddress(int col, int row)
        {
            string colStr = "";
            int tempCol = col;
            while (tempCol > 0)
            {
                int rem = (tempCol - 1) % 26;
                colStr = (char)('A' + rem) + colStr;
                tempCol = (tempCol - 1) / 26;
            }
            return $"{colStr}{row}";
        }

        private object EvaluateBinary(BinaryOpNode node)
        {
            var left = Evaluate(node.Left);
            if (left is CalcError) return left;

            var right = Evaluate(node.Right);
            if (right is CalcError) return right;

            if (node.Op == TokenType.Ampersand)
            {
                return Convert.ToString(left) + Convert.ToString(right);
            }

            // Numeric operations
            if (!TryConvertToDouble(left, out double l)) return CalcError.Value;
            if (!TryConvertToDouble(right, out double r)) return CalcError.Value;

            switch (node.Op)
            {
                case TokenType.Plus: return l + r;
                case TokenType.Minus: return l - r;
                case TokenType.Multiply: return l * r;
                case TokenType.Divide: 
                    if (r == 0) return CalcError.Div0;
                    return l / r;
                case TokenType.Power: return Math.Pow(l, r);
                // Comparison
                case TokenType.Equal: return Compare(left, right) == 0;
                case TokenType.NotEqual: return Compare(left, right) != 0;
                case TokenType.LessThan: return Compare(left, right) < 0;
                case TokenType.GreaterThan: return Compare(left, right) > 0;
                case TokenType.LessThanOrEqual: return Compare(left, right) <= 0;
                case TokenType.GreaterThanOrEqual: return Compare(left, right) >= 0;
                default: return CalcError.Value;
            }
        }

        private object EvaluateUnary(UnaryOpNode node)
        {
            var val = Evaluate(node.Operand);
            if (val is CalcError) return val;

            if (!TryConvertToDouble(val, out double d)) return CalcError.Value;

            if (node.Op == TokenType.Minus) return -d;
            return d;
        }

        private object EvaluateFunction(FunctionCallNode node)
        {
            var args = new object[node.Arguments.Count];
            for (int i = 0; i < node.Arguments.Count; i++)
            {
                args[i] = Evaluate(node.Arguments[i]);
                // Note: We don't return immediately on error here because some functions (like ISERROR) might handle errors.
                // But for standard math functions, they will check args and return error.
            }
            
            try
            {
                return _registry.Call(node.FunctionName, args);
            }
            catch (KeyNotFoundException)
            {
                return CalcError.Name;
            }
            catch (Exception)
            {
                return CalcError.Value;
            }
        }

        public static bool TryConvertToDouble(object val, out double result)
        {
            result = 0;
            if (val == null) return true; // null -> 0
            if (val is CalcError) return false;
            if (val is double d) { result = d; return true; }
            if (val is int i) { result = i; return true; }
            if (val is bool b) { result = b ? 1.0 : 0.0; return true; }
            if (val is string s)
            {
                return double.TryParse(s, out result);
            }
            try
            {
                result = Convert.ToDouble(val);
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        private int Compare(object left, object right)
        {
            // Try numeric comparison first
            if (TryConvertToDouble(left, out double l) && TryConvertToDouble(right, out double r))
            {
                return l.CompareTo(r);
            }

            // Fallback to string comparison
            string sL = Convert.ToString(left) ?? "";
            string sR = Convert.ToString(right) ?? "";
            return string.Compare(sL, sR, StringComparison.OrdinalIgnoreCase);
        }
    }
}
