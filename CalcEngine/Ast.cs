using System.Collections.Generic;

namespace CalcEngine
{
    public abstract class AstNode { }

    public class NumberNode : AstNode
    {
        public double Value { get; }
        public NumberNode(double value) { Value = value; }
    }

    public class StringNode : AstNode
    {
        public string Value { get; }
        public StringNode(string value) { Value = value; }
    }

    public class CellReferenceNode : AstNode
    {
        public string Address { get; }
        public CellReferenceNode(string address) { Address = address; }
    }

    public class RangeNode : AstNode
    {
        public string StartAddress { get; }
        public string EndAddress { get; }
        public RangeNode(string start, string end) { StartAddress = start; EndAddress = end; }
    }

    public class BinaryOpNode : AstNode
    {
        public AstNode Left { get; }
        public AstNode Right { get; }
        public TokenType Op { get; }
        public BinaryOpNode(AstNode left, TokenType op, AstNode right)
        {
            Left = left;
            Op = op;
            Right = right;
        }
    }

    public class UnaryOpNode : AstNode
    {
        public TokenType Op { get; }
        public AstNode Operand { get; }
        public UnaryOpNode(TokenType op, AstNode operand)
        {
            Op = op;
            Operand = operand;
        }
    }

    public class FunctionCallNode : AstNode
    {
        public string FunctionName { get; }
        public List<AstNode> Arguments { get; }
        public FunctionCallNode(string name, List<AstNode> args)
        {
            FunctionName = name;
            Arguments = args;
        }
    }
}
