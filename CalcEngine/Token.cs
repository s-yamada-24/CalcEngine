namespace CalcEngine
{
    public enum TokenType
    {
        Number,
        String,
        Identifier, // Cell reference or function name
        Plus,
        Minus,
        Multiply,
        Divide,
        Power,
        LParen,
        RParen,
        Comma,
        Colon, // For ranges
        EOF,
        // Comparison
        Equal,
        NotEqual,
        LessThan,
        GreaterThan,
        LessThanOrEqual,
        GreaterThanOrEqual,
        Ampersand // String concatenation
    }

    public class Token
    {
        public TokenType Type { get; }
        public string Value { get; }
        public int Position { get; }

        public Token(TokenType type, string value, int position)
        {
            Type = type;
            Value = value;
            Position = position;
        }

        public override string ToString() => $"{Type}: {Value}";
    }
}
