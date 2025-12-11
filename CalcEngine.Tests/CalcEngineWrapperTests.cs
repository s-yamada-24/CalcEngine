using Xunit;
using CalcEngine;

namespace CalcEngine.Tests
{
    public class CalcEngineWrapperTests
    {
        [Fact]
        public void TestCalcEngineWrapper()
        {
            // Initialize the engine
            var engine = new CalcEngine();

            // Set values
            engine.SetValue("A1", 10);
            engine.SetValue("A2", 20);

            // Evaluate formula
            var result = engine.Evaluate("=A1 + A2");

            // Assert result
            Assert.Equal(30.0, result);
        }

        [Fact]
        public void TestCustomFunctionRegistration()
        {
            var engine = new CalcEngine();
            engine.SetValue("A1", 5);

            // Register custom function
            engine.RegisterFunction("TRIPLE", args => 
            {
                return System.Convert.ToDouble(args[0]) * 3;
            });

            var result = engine.Evaluate("=TRIPLE(A1)");
            Assert.Equal(15.0, result);
        }
    }
}
