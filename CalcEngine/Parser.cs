using System;
using System.Collections.Generic;

namespace CalcEngine
{
    public class Parser
    {
        private readonly List<Token> _tokens;
        private int _position;

        public Parser(List<Token> tokens)
        {
            _tokens = tokens;
            _position = 0;
        }

        public AstNode Parse()
        {
            var node = ParseExpression();
            if (Current.Type != TokenType.EOF)
                throw new FormatException($"Unexpected token {Current.Type} at position {Current.Position}");
            return node;
        }

        private Token Current => _position < _tokens.Count ? _tokens[_position] : _tokens[_tokens.Count - 1];

        private Token Consume(TokenType type)
        {
            if (Current.Type == type)
            {
                var token = Current;
                _position++;
                return token;
            }
            throw new FormatException($"Expected {type} but found {Current.Type} at position {Current.Position}");
        }

        // Expression -> Comparison
        private AstNode ParseExpression()
        {
            return ParseComparison();
        }

        // Comparison -> Concat ( (= | <> | < | > | <= | >=) Concat )*
        private AstNode ParseComparison()
        {
            var left = ParseConcat();

            while (Current.Type == TokenType.Equal || Current.Type == TokenType.NotEqual ||
                   Current.Type == TokenType.LessThan || Current.Type == TokenType.GreaterThan ||
                   Current.Type == TokenType.LessThanOrEqual || Current.Type == TokenType.GreaterThanOrEqual)
            {
                var op = Current.Type;
                _position++;
                var right = ParseConcat();
                left = new BinaryOpNode(left, op, right);
            }
            return left;
        }

        // Concat -> AddSub ( & AddSub )*
        private AstNode ParseConcat()
        {
            var left = ParseAddSub();
            while (Current.Type == TokenType.Ampersand)
            {
                _position++;
                var right = ParseAddSub();
                left = new BinaryOpNode(left, TokenType.Ampersand, right);
            }
            return left;
        }

        // AddSub -> MulDiv ( (+|-) MulDiv )*
        private AstNode ParseAddSub()
        {
            var left = ParseMulDiv();
            while (Current.Type == TokenType.Plus || Current.Type == TokenType.Minus)
            {
                var op = Current.Type;
                _position++;
                var right = ParseMulDiv();
                left = new BinaryOpNode(left, op, right);
            }
            return left;
        }

        // MulDiv -> Power ( (*|/) Power )*
        private AstNode ParseMulDiv()
        {
            var left = ParsePower();
            while (Current.Type == TokenType.Multiply || Current.Type == TokenType.Divide)
            {
                var op = Current.Type;
                _position++;
                var right = ParsePower();
                left = new BinaryOpNode(left, op, right);
            }
            return left;
        }

        // Power -> Unary ( ^ Unary )*
        private AstNode ParsePower()
        {
            var left = ParseUnary();
            while (Current.Type == TokenType.Power)
            {
                _position++;
                var right = ParseUnary();
                left = new BinaryOpNode(left, TokenType.Power, right);
            }
            return left;
        }

        // Unary -> (+|-) Unary | Primary
        private AstNode ParseUnary()
        {
            if (Current.Type == TokenType.Plus || Current.Type == TokenType.Minus)
            {
                var op = Current.Type;
                _position++;
                var operand = ParseUnary();
                return new UnaryOpNode(op, operand);
            }
            return ParsePrimary();
        }

        // Primary -> Number | String | Identifier (Function or Reference) | ( Expression )
        private AstNode ParsePrimary()
        {
            if (Current.Type == TokenType.Number)
            {
                var value = double.Parse(Current.Value);
                _position++;
                return new NumberNode(value);
            }

            if (Current.Type == TokenType.String)
            {
                var value = Current.Value;
                _position++;
                return new StringNode(value);
            }

            if (Current.Type == TokenType.Identifier)
            {
                var name = Current.Value;
                _position++;

                // Function Call: Identifier ( ... )
                if (Current.Type == TokenType.LParen)
                {
                    _position++; // Consume '('
                    var args = new List<AstNode>();
                    if (Current.Type != TokenType.RParen)
                    {
                        do
                        {
                            args.Add(ParseExpression());
                        } while (Current.Type == TokenType.Comma && _position++ > 0); // Consume comma
                    }
                    Consume(TokenType.RParen);
                    return new FunctionCallNode(name, args);
                }
                
                // Range: Identifier : Identifier
                if (Current.Type == TokenType.Colon)
                {
                    _position++; // Consume ':'
                    if (Current.Type != TokenType.Identifier)
                        throw new FormatException($"Expected identifier after colon at {Current.Position}");
                    var endName = Current.Value;
                    _position++;
                    return new RangeNode(name, endName);
                }

                return new CellReferenceNode(name);
            }

            if (Current.Type == TokenType.LParen)
            {
                _position++;
                var node = ParseExpression();
                Consume(TokenType.RParen);
                return node;
            }

            throw new FormatException($"Unexpected token {Current.Type} at position {Current.Position}");
        }
    }
}
