using Xunit;
using CalcEngine;
using System.Collections.Generic;

namespace CalcEngine.Tests
{
    public class FormulaEvaluatorTests
    {
        [Fact]
        public void TestBasicArithmetic()
        {
            var table = new VirtualTable();
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal(3.0, evaluator.Evaluate("=1 + 2"));
            Assert.Equal(7.0, evaluator.Evaluate("=1 + 2 * 3"));
            Assert.Equal(9.0, evaluator.Evaluate("=(1 + 2) * 3"));
            Assert.Equal(2.5, evaluator.Evaluate("=5 / 2"));
        }

        [Fact]
        public void TestCellReferences()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 10);
            table.SetValue("B1", 20);
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal(30.0, evaluator.Evaluate("=A1 + B1"));
            Assert.Equal(200.0, evaluator.Evaluate("=A1 * B1"));
        }

        [Fact]
        public void TestMissingCellReference()
        {
            var table = new VirtualTable();
            var evaluator = new FormulaEvaluator(table);

            // Missing cell should return #REF!
            var result = evaluator.Evaluate("=A1");
            Assert.IsType<CalcError>(result);
            Assert.Equal(CalcError.Ref.Code, result.ToString());

            // Error propagation
            var result2 = evaluator.Evaluate("=A1 + 1");
            Assert.IsType<CalcError>(result2);
            Assert.Equal(CalcError.Ref.Code, result2.ToString());
        }

        [Fact]
        public void TestDivisionByZero()
        {
            var table = new VirtualTable();
            var evaluator = new FormulaEvaluator(table);

            var result = evaluator.Evaluate("=1 / 0");
            Assert.IsType<CalcError>(result);
            Assert.Equal(CalcError.Div0.Code, result.ToString());
        }

        [Fact]
        public void TestFunctions()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 10);
            table.SetValue("A2", 20);
            table.SetValue("A3", 30);
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal(60.0, evaluator.Evaluate("=SUM(A1:A3)"));
            Assert.Equal(20.0, evaluator.Evaluate("=AVERAGE(A1:A3)"));
            Assert.Equal(30.0, evaluator.Evaluate("=MAX(A1:A3)"));
            Assert.Equal(10.0, evaluator.Evaluate("=MIN(A1:A3)"));
            Assert.Equal(3.0, evaluator.Evaluate("=COUNT(A1:A3)"));
        }

        [Fact]
        public void TestNestedParentheses()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 1);
            table.SetValue("B1", 2);
            table.SetValue("C1", 3);
            table.SetValue("D1", 4);
            var evaluator = new FormulaEvaluator(table);

            // ((1+2) * (3-4)) = 3 * -1 = -3
            Assert.Equal(-3.0, evaluator.Evaluate("=((A1 + B1) * (C1 - D1))"));
        }

        [Fact]
        public void TestStringConcatenation()
        {
            var table = new VirtualTable();
            table.SetValue("A1", "Hello");
            table.SetValue("B1", "World");
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal("HelloWorld", evaluator.Evaluate("=A1 & B1"));
            Assert.Equal("Hello World", evaluator.Evaluate("=A1 & \" \" & B1"));
        }

        [Fact]
        public void TestLogic()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 10);
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal("Big", evaluator.Evaluate("=IF(A1 > 5, \"Big\", \"Small\")"));
            Assert.Equal("Small", evaluator.Evaluate("=IF(A1 < 5, \"Big\", \"Small\")"));
            Assert.Equal(true, evaluator.Evaluate("=AND(A1 > 5, A1 < 20)"));
        }
        
        [Fact]
        public void TestCountIfSumIf()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 10);
            table.SetValue("A2", 20);
            table.SetValue("A3", 10);
            var evaluator = new FormulaEvaluator(table);

            Assert.Equal(2.0, evaluator.Evaluate("=COUNTIF(A1:A3, 10)"));
            Assert.Equal(1.0, evaluator.Evaluate("=COUNTIF(A1:A3, \">15\")"));
            Assert.Equal(20.0, evaluator.Evaluate("=SUMIF(A1:A3, 10)")); // 10 + 10 = 20
        }
    }
}
