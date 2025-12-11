using System;
using System.Collections.Generic;
using System.Linq;

namespace CalcEngine
{
    public static class StandardFunctions
    {
        public static void Register(FunctionRegistry registry)
        {
            // Math
            registry.Register("SUM", Sum);
            registry.Register("AVERAGE", Average);
            registry.Register("MIN", Min);
            registry.Register("MAX", Max);
            registry.Register("COUNT", Count);
            registry.Register("COUNTIF", CountIf);
            registry.Register("SUMIF", SumIf);
            registry.Register("ROUND", Round);
            registry.Register("ABS", Abs);
            registry.Register("SQRT", Sqrt);

            // Logic
            registry.Register("IF", If);
            registry.Register("AND", And);
            registry.Register("OR", Or);
            registry.Register("NOT", Not);

            // String
            registry.Register("CONCATENATE", Concatenate);
            registry.Register("LEFT", Left);
            registry.Register("RIGHT", Right);
            registry.Register("MID", Mid);
            registry.Register("LEN", Len);
            registry.Register("UPPER", Upper);
            registry.Register("LOWER", Lower);
        }

        private static CalcError? CheckErrors(object[] args)
        {
            foreach (var arg in args)
            {
                if (arg is CalcError err) return err;
            }
            return null;
        }

        private static IEnumerable<double> FlattenNumbers(object[] args)
        {
            foreach (var arg in args)
            {
                if (arg is IEnumerable<object> list)
                {
                    foreach (var item in list)
                    {
                        if (Evaluator.TryConvertToDouble(item, out double d)) yield return d;
                    }
                }
                else if (Evaluator.TryConvertToDouble(arg, out double d))
                {
                    yield return d;
                }
            }
        }

        private static object Sum(object[] args)
        {
            // SUM ignores text and errors in ranges usually, but if direct arg is error, it returns error.
            // For simplicity, let's propagate errors if they are direct args.
            // But for ranges, we filter.
            return FlattenNumbers(args).Sum();
        }

        private static object Average(object[] args)
        {
            var nums = FlattenNumbers(args).ToList();
            if (nums.Count == 0) return CalcError.Div0;
            return nums.Average();
        }

        private static object Min(object[] args)
        {
            var nums = FlattenNumbers(args).ToList();
            if (nums.Count == 0) return 0.0;
            return nums.Min();
        }

        private static object Max(object[] args)
        {
            var nums = FlattenNumbers(args).ToList();
            if (nums.Count == 0) return 0.0;
            return nums.Max();
        }

        private static object Count(object[] args)
        {
            return (double)FlattenNumbers(args).Count();
        }

        private static object CountIf(object[] args)
        {
            if (args.Length != 2) return CalcError.Value;
            var range = args[0] as IEnumerable<object>;
            if (range == null) range = new[] { args[0] }; 
            
            var criteria = args[1];
            int count = 0;
            foreach (var item in range)
            {
                // Skip errors in range for COUNTIF? Excel usually ignores errors in range unless criteria matches error.
                if (item is CalcError) continue;
                if (EvaluateCriteria(item, criteria)) count++;
            }
            return (double)count;
        }

        private static object SumIf(object[] args)
        {
            if (args.Length < 2 || args.Length > 3) return CalcError.Value;
            var range = (args[0] as IEnumerable<object>)?.ToList();
            if (range == null) range = new List<object> { args[0] };
            
            var criteria = args[1];
            var sumRange = (args.Length == 3 ? args[2] as IEnumerable<object> : range)?.ToList();
            if (sumRange == null) sumRange = new List<object> { args.Length == 3 ? args[2] : args[0] };

            double sum = 0;
            for (int i = 0; i < range.Count; i++)
            {
                if (i >= sumRange.Count) break;
                if (range[i] is CalcError) continue;

                if (EvaluateCriteria(range[i], criteria))
                {
                    if (Evaluator.TryConvertToDouble(sumRange[i], out double d)) sum += d;
                }
            }
            return sum;
        }

        private static bool EvaluateCriteria(object item, object criteria)
        {
            string cStr = criteria?.ToString() ?? "";
            if (cStr.StartsWith(">") || cStr.StartsWith("<") || cStr.StartsWith("=") || cStr.StartsWith("<>") || cStr.StartsWith(">=") || cStr.StartsWith("<="))
            {
                string op = "";
                if (cStr.StartsWith(">=") || cStr.StartsWith("<=") || cStr.StartsWith("<>")) { op = cStr.Substring(0, 2); cStr = cStr.Substring(2); }
                else { op = cStr.Substring(0, 1); cStr = cStr.Substring(1); }
                
                if (!double.TryParse(cStr, out double cVal)) return false; 
                if (!Evaluator.TryConvertToDouble(item, out double iVal)) return false;

                switch (op)
                {
                    case ">": return iVal > cVal;
                    case "<": return iVal < cVal;
                    case "=": return Math.Abs(iVal - cVal) < 1e-9;
                    case "<>": return Math.Abs(iVal - cVal) > 1e-9;
                    case ">=": return iVal >= cVal;
                    case "<=": return iVal <= cVal;
                }
            }
            
            return item?.ToString() == criteria?.ToString();
        }

        private static object Round(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 2) return CalcError.Value;
            if (!Evaluator.TryConvertToDouble(args[0], out double num)) return CalcError.Value;
            if (!Evaluator.TryConvertToDouble(args[1], out double digits)) return CalcError.Value;
            return Math.Round(num, (int)digits, MidpointRounding.AwayFromZero);
        }

        private static object Abs(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            if (!Evaluator.TryConvertToDouble(args[0], out double val)) return CalcError.Value;
            return Math.Abs(val);
        }

        private static object Sqrt(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            if (!Evaluator.TryConvertToDouble(args[0], out double val)) return CalcError.Value;
            if (val < 0) return CalcError.Num;
            return Math.Sqrt(val);
        }

        private static object If(object[] args)
        {
            // IF evaluates condition first.
            if (args.Length != 3) return CalcError.Value;
            if (args[0] is CalcError err) return err;
            
            // Excel treats 0 as false, others as true for numbers.
            bool cond = false;
            if (args[0] is bool b) cond = b;
            else if (Evaluator.TryConvertToDouble(args[0], out double d)) cond = d != 0;
            else return CalcError.Value;

            return cond ? args[1] : args[2];
        }

        private static object And(object[] args)
        {
            foreach (var arg in args)
            {
                if (arg is CalcError err) return err;
                bool b = false;
                if (arg is bool bl) b = bl;
                else if (Evaluator.TryConvertToDouble(arg, out double d)) b = d != 0;
                else return CalcError.Value;
                
                if (!b) return false;
            }
            return true;
        }

        private static object Or(object[] args)
        {
            foreach (var arg in args)
            {
                if (arg is CalcError err) return err;
                bool b = false;
                if (arg is bool bl) b = bl;
                else if (Evaluator.TryConvertToDouble(arg, out double d)) b = d != 0;
                else return CalcError.Value;

                if (b) return true;
            }
            return false;
        }

        private static object Not(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            
            bool b = false;
            if (args[0] is bool bl) b = bl;
            else if (Evaluator.TryConvertToDouble(args[0], out double d)) b = d != 0;
            else return CalcError.Value;

            return !b;
        }

        private static object Concatenate(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            return string.Concat(args);
        }

        private static object Left(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 2) return CalcError.Value;
            string s = args[0].ToString();
            if (!Evaluator.TryConvertToDouble(args[1], out double lenD)) return CalcError.Value;
            int len = (int)lenD;
            if (len < 0) return CalcError.Value;
            if (len > s.Length) len = s.Length;
            return s.Substring(0, len);
        }

        private static object Right(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 2) return CalcError.Value;
            string s = args[0].ToString();
            if (!Evaluator.TryConvertToDouble(args[1], out double lenD)) return CalcError.Value;
            int len = (int)lenD;
            if (len < 0) return CalcError.Value;
            if (len > s.Length) len = s.Length;
            return s.Substring(s.Length - len);
        }

        private static object Mid(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 3) return CalcError.Value;
            string s = args[0].ToString();
            if (!Evaluator.TryConvertToDouble(args[1], out double startD)) return CalcError.Value;
            if (!Evaluator.TryConvertToDouble(args[2], out double lenD)) return CalcError.Value;
            
            int start = (int)startD - 1; 
            int len = (int)lenD;
            
            if (start < 0) start = 0; // Excel behavior might differ but this is safe
            if (start >= s.Length) return "";
            if (start + len > s.Length) len = s.Length - start;
            return s.Substring(start, len);
        }

        private static object Len(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            return (double)args[0].ToString().Length;
        }

        private static object Upper(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            return args[0].ToString().ToUpper();
        }

        private static object Lower(object[] args)
        {
            if (CheckErrors(args) is CalcError err) return err;
            if (args.Length != 1) return CalcError.Value;
            return args[0].ToString().ToLower();
        }
    }
}
