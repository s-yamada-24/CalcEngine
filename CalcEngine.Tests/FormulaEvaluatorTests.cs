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

            // Missing cell should return empty string
            var result = evaluator.Evaluate("=A1");
            Assert.Equal("", result);

            // In numeric context, empty string is treated as 0
            var result2 = evaluator.Evaluate("=A1 + 1");
            Assert.Equal(1.0, result2);
            
            // In string context, empty string remains empty
            var result3 = evaluator.Evaluate("=A1 & \"test\"");
            Assert.Equal("test", result3);
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

        [Fact]
        public void TestModAndTrim()
        {
            var table = new VirtualTable();
            table.SetValue("A1", "  Hello World  ");
            table.SetValue("N1", 3);
            table.SetValue("N2", 2);
            table.SetValue("N3", -3);
            table.SetValue("N4", -2);
            var evaluator = new FormulaEvaluator(table);

            // TRIM
            Assert.Equal("Hello World", evaluator.Evaluate("=TRIM(A1)"));

            // MOD
            // 3 % 2 = 1
            Assert.Equal(1.0, evaluator.Evaluate("=MOD(N1, N2)"));
            // -3 % 2 = 1 (Excel behavior)
            Assert.Equal(1.0, evaluator.Evaluate("=MOD(N3, N2)"));
            // 3 % -2 = -1 (Excel behavior)
            Assert.Equal(-1.0, evaluator.Evaluate("=MOD(N1, N4)"));
            // -3 % -2 = -1
            Assert.Equal(-1.0, evaluator.Evaluate("=MOD(N3, N4)"));
            
            // MOD zero division
            var err = evaluator.Evaluate("=MOD(N1, 0)");
            Assert.IsType<CalcError>(err);
            Assert.Equal(CalcError.Div0.Code, err.ToString());
        }

        [Fact]
        public void TestStringFunctionsWithEmptyString()
        {
            var table = new VirtualTable();
            table.SetValue("A1", "");
            table.SetValue("A2", "Hello");
            var evaluator = new FormulaEvaluator(table);

            // LEFT with empty string
            Assert.Equal("", evaluator.Evaluate("=LEFT(A1, 5)"));
            
            // RIGHT with empty string
            Assert.Equal("", evaluator.Evaluate("=RIGHT(A1, 3)"));
            
            // MID with empty string
            Assert.Equal("", evaluator.Evaluate("=MID(A1, 1, 5)"));
            
            // LEN with empty string
            Assert.Equal(0.0, evaluator.Evaluate("=LEN(A1)"));
            
            // UPPER with empty string
            Assert.Equal("", evaluator.Evaluate("=UPPER(A1)"));
            
            // LOWER with empty string
            Assert.Equal("", evaluator.Evaluate("=LOWER(A1)"));
            
            // TRIM with empty string
            Assert.Equal("", evaluator.Evaluate("=TRIM(A1)"));
            
            // CONCATENATE with empty string
            Assert.Equal("Hello", evaluator.Evaluate("=CONCATENATE(A1, A2)"));
        }

        [Fact]
        public void TestStringFunctionsWithMissingCell()
        {
            var table = new VirtualTable();
            var evaluator = new FormulaEvaluator(table);

            // Missing cell in string functions should be treated as empty string
            Assert.Equal("", evaluator.Evaluate("=LEFT(B1, 5)"));
            Assert.Equal("", evaluator.Evaluate("=RIGHT(B1, 3)"));
            Assert.Equal("", evaluator.Evaluate("=MID(B1, 1, 5)"));
            Assert.Equal(0.0, evaluator.Evaluate("=LEN(B1)"));
            Assert.Equal("", evaluator.Evaluate("=UPPER(B1)"));
            Assert.Equal("", evaluator.Evaluate("=LOWER(B1)"));
            Assert.Equal("", evaluator.Evaluate("=TRIM(B1)"));
        }

        [Fact]
        public void TestSumWithMissingCells()
        {
            var table = new VirtualTable();
            table.SetValue("A1", 10);
            // A2 is missing
            table.SetValue("A3", 20);
            var evaluator = new FormulaEvaluator(table);

            // SUM should treat missing cells as 0
            Assert.Equal(30.0, evaluator.Evaluate("=SUM(A1:A3)"));
        }
    }
}
