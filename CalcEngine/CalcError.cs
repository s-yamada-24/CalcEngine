namespace CalcEngine
{
    public class CalcError
    {
        public string Code { get; }

        private CalcError(string code)
        {
            Code = code;
        }

        public override string ToString() => Code;

        public static readonly CalcError Div0 = new("#DIV/0!");
        public static readonly CalcError Value = new("#VALUE!");
        public static readonly CalcError Ref = new("#REF!");
        public static readonly CalcError Name = new("#NAME?");
        public static readonly CalcError NA = new("#N/A");
        public static readonly CalcError Num = new("#NUM!");
    }
}
