using System;
using System.Collections.Generic;
using System.Text;

namespace CalcEngine
{
    public class Lexer
    {
        private readonly string _input;
        private int _position;

        public Lexer(string input)
        {
            _input = input;
            _position = 0;
        }

        public List<Token> Tokenize()
        {
            var tokens = new List<Token>();
            while (_position < _input.Length)
            {
                char current = _input[_position];

                if (char.IsWhiteSpace(current))
                {
                    _position++;
                    continue;
                }

                if (char.IsDigit(current))
                {
                    tokens.Add(ReadNumber());
                    continue;
                }

                if (char.IsLetter(current) || current == '_')
                {
                    tokens.Add(ReadIdentifier());
                    continue;
                }

                if (current == '"')
                {
                    tokens.Add(ReadString());
                    continue;
                }

                switch (current)
                {
                    case '+': tokens.Add(new Token(TokenType.Plus, "+", _position++)); break;
                    case '-': tokens.Add(new Token(TokenType.Minus, "-", _position++)); break;
                    case '*': tokens.Add(new Token(TokenType.Multiply, "*", _position++)); break;
                    case '/': tokens.Add(new Token(TokenType.Divide, "/", _position++)); break;
                    case '^': tokens.Add(new Token(TokenType.Power, "^", _position++)); break;
                    case '(': tokens.Add(new Token(TokenType.LParen, "(", _position++)); break;
                    case ')': tokens.Add(new Token(TokenType.RParen, ")", _position++)); break;
                    case ',': tokens.Add(new Token(TokenType.Comma, ",", _position++)); break;
                    case ':': tokens.Add(new Token(TokenType.Colon, ":", _position++)); break;
                    case '&': tokens.Add(new Token(TokenType.Ampersand, "&", _position++)); break;
                    case '=': tokens.Add(new Token(TokenType.Equal, "=", _position++)); break;
                    case '<':
                        if (Peek() == '>') { tokens.Add(new Token(TokenType.NotEqual, "<>", _position)); _position += 2; }
                        else if (Peek() == '=') { tokens.Add(new Token(TokenType.LessThanOrEqual, "<=", _position)); _position += 2; }
                        else { tokens.Add(new Token(TokenType.LessThan, "<", _position++)); }
                        break;
                    case '>':
                        if (Peek() == '=') { tokens.Add(new Token(TokenType.GreaterThanOrEqual, ">=", _position)); _position += 2; }
                        else { tokens.Add(new Token(TokenType.GreaterThan, ">", _position++)); }
                        break;
                    default:
                        throw new FormatException($"Unexpected character '{current}' at position {_position}");
                }
            }

            tokens.Add(new Token(TokenType.EOF, "", _position));
            return tokens;
        }

        private char Peek()
        {
            if (_position + 1 >= _input.Length) return '\0';
            return _input[_position + 1];
        }

        private Token ReadNumber()
        {
            int start = _position;
            while (_position < _input.Length && (char.IsDigit(_input[_position]) || _input[_position] == '.'))
            {
                _position++;
            }
            return new Token(TokenType.Number, _input.Substring(start, _position - start), start);
        }

        private Token ReadIdentifier()
        {
            int start = _position;
            while (_position < _input.Length && (char.IsLetterOrDigit(_input[_position]) || _input[_position] == '_'))
            {
                _position++;
            }
            return new Token(TokenType.Identifier, _input.Substring(start, _position - start), start);
        }

        private Token ReadString()
        {
            int start = _position;
            _position++; // Skip opening quote
            var sb = new StringBuilder();
            while (_position < _input.Length)
            {
                if (_input[_position] == '"')
                {
                    if (Peek() == '"') // Escaped quote
                    {
                        sb.Append('"');
                        _position += 2;
                    }
                    else
                    {
                        _position++; // Closing quote
                        return new Token(TokenType.String, sb.ToString(), start);
                    }
                }
                else
                {
                    sb.Append(_input[_position]);
                    _position++;
                }
            }
            throw new FormatException("Unterminated string literal");
        }
    }
}
