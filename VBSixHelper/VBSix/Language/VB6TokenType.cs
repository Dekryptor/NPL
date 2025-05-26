namespace VBSix.Language
{
    /// <summary>
    /// Defines the various types of tokens that can be encountered
    /// while parsing Visual Basic 6 source code.
    /// </summary>
    public enum VB6TokenType
    {
        Unknown,        // For unrecognized sequences

        // Basic Structure
        Whitespace,     // Spaces, tabs
        Newline,        // CR, LF, CRLF
        Comment,        // Starts with ' or REM

        // Literals
        LiteralString,  // e.g., "Hello World"
        LiteralNumber,  // e.g., 123, 3.14, &HFF, &O77
        // LiteralDate, // e.g., #1/1/2000# (might be parsed as Octothorpe, Number, Slash, etc. initially)

        // Identifiers
        Identifier,     // Variable names, function names, user-defined types, etc.

        // Keywords (Alphabetical Grouping for Readability)
        // A
        AddressOfKeyword,
        AliasKeyword,
        AndKeyword,
        AppActivateKeyword,
        AsKeyword,
        // B
        BaseKeyword, // Option Base
        BeepKeyword,
        BinaryKeyword, // Option Compare Binary
        BooleanKeyword,
        ByRefKeyword,
        ByValKeyword,
        ByteKeyword,
        // C
        CallKeyword,
        CaseKeyword,
        ChDirKeyword,
        ChDriveKeyword,
        CloseKeyword,
        CompareKeyword, // Option Compare
        ConstKeyword,
        CurrencyKeyword,
        // D
        DateKeyword,
        DecimalKeyword,
        DeclareKeyword,
        DefBoolKeyword,
        DefByteKeyword,
        DefCurKeyword,
        DefDateKeyword,
        DefDblKeyword,
        DefDecKeyword,
        DefIntKeyword,
        DefLngKeyword,
        DefObjKeyword,
        DefSngKeyword,
        DefStrKeyword,
        DefVarKeyword,
        DeleteSettingKeyword,
        DimKeyword,
        DoKeyword,
        DoubleKeyword,
        // E
        ElseKeyword,
        ElseIfKeyword,
        EmptyKeyword,
        EndKeyword,
        EnumKeyword,
        EqvKeyword,
        EraseKeyword,
        ErrorKeyword,
        EventKeyword,
        ExitKeyword,
        ExplicitKeyword, // Option Explicit
        // F
        FalseKeyword,
        FileCopyKeyword,
        ForKeyword,
        FriendKeyword,
        FunctionKeyword,
        // G
        GetKeyword, // Property Get
        GotoKeyword,
        // I
        IfKeyword,
        ImpKeyword,
        ImplementsKeyword,
        InputKeyword, // For Input #
        IntegerKeyword,
        IsKeyword,
        // K
        KillKeyword,
        // L
        LenKeyword, // Often a function, but can appear as keyword context
        LetKeyword, // Property Let
        LibKeyword,
        LikeKeyword,
        LineKeyword, // Line Input #, or Line statement
        LoadKeyword,
        LockKeyword,
        LongKeyword,
        LSetKeyword,
        // M
        MeKeyword,
        MidKeyword, // Often a function
        MkDirKeyword,
        ModKeyword,
        // N
        NameKeyword, // File Name statement
        NewKeyword,
        NextKeyword,
        NotKeyword,
        NullKeyword,
        // O
        ObjectKeyword,
        OnKeyword,
        OpenKeyword,
        OptionKeyword,
        OptionalKeyword,
        OrKeyword,
        // P
        ParamArrayKeyword,
        PreserveKeyword,
        PrintKeyword,
        PrivateKeyword,
        PropertyKeyword, // Property Get/Let/Set
        PublicKeyword,
        PutKeyword,
        // R
        RaiseEventKeyword,
        RandomizeKeyword,
        ReDimKeyword,
        ResetKeyword,
        ResumeKeyword,
        RmDirKeyword,
        RSetKeyword,
        // S
        SavePictureKeyword,
        SaveSettingKeyword,
        SeekKeyword,
        SelectKeyword, // Select Case
        SendKeysKeyword,
        SetKeyword, // Property Set, or Set object =
        SetAttrKeyword,
        SingleKeyword,
        StaticKeyword,
        StepKeyword,
        StopKeyword,
        StringKeyword, // Data type
        SubKeyword,
        // T
        ThenKeyword,
        TimeKeyword,
        ToKeyword,
        TrueKeyword,
        TypeKeyword, // User-defined type
        // U
        UnlockKeyword,
        UntilKeyword,
        // V
        VariantKeyword,
        // W
        WendKeyword,
        WhileKeyword,
        WidthKeyword, // For Print #, Open
        WithKeyword,
        WithEventsKeyword,
        WriteKeyword, // Write #
        // X
        XorKeyword,

        // Symbols and Operators
        EqualityOperator,       // = (also assignment)
        NotEqualOperator,       // <>
        LessThanOperator,       // <
        GreaterThanOperator,    // >
        LessThanOrEqualOperator, // <=
        GreaterThanOrEqualOperator, // >=
        AssignmentOperator,     // = (specific context if distinguished from equality) - usually just EqualityOperator
        
        LeftParenthesis,        // (
        RightParenthesis,       // )
        LeftSquareBracket,      // [ (used for some attributes or object library names)
        RightSquareBracket,     // ]
        
        Comma,                  // ,
        Semicolon,              // ;
        Colon,                  // : (statement separator)
        Dot,                    // . (member access)
        
        PlusOperator,           // +
        MinusOperator,          // -
        MultiplyOperator,       // *
        DivideOperator,         // / (floating point division)
        IntegerDivideOperator,  // \ (integer division)
        ExponentiateOperator,   // ^
        
        Ampersand,              // & (string concatenation, type declaration for Long)
        Underscore,             // _ (line continuation)

        // Type Suffixes / Special Characters
        DollarSign,             // $ (String type suffix)
        Percent,                // % (Integer type suffix)
        Octothorpe,             // # (Double type suffix, Date literal delimiter)
        ExclamationMark,        // ! (Single type suffix, Dictionary item access)
        AtSign,                 // @ (Currency type suffix)

        // Punctuation might be redundant if specific operators/symbols are used.
        // Keeping it general for now, can be refined.
        Punctuation,
        LineContinuation
    }

    /// <summary>
    /// Represents a single token parsed from VB6 source code.
    /// It includes the token type, its string value, and positional information.
    /// </summary>
    public class VB6Token
    {
        /// <summary>
        /// The type of the token.
        /// </summary>
        public VB6TokenType Type { get; }

        /// <summary>
        /// The string value or representation of the token.
        /// For keywords, this is the keyword itself (e.g., "Dim").
        /// For identifiers, it's the name (e.g., "myVariable").
        /// For literals, it's the literal value (e.g., "123", "\"Hello\"").
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// The original byte sequence from the source file that formed this token.
        /// Useful for debugging, exact source reproduction, or handling encoding nuances.
        /// </summary>
        public byte[] OriginalBytes { get; } 

        /// <summary>
        /// The 1-based line number where the token begins.
        /// </summary>
        public int LineNumber { get; }

        /// <summary>
        /// The 1-based column number where the token begins on its line.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VB6Token"/> class.
        /// </summary>
        /// <param name="type">The type of the token.</param>
        /// <param name="value">The string value of the token.</param>
        /// <param name="lineNumber">The line number where the token starts.</param>
        /// <param name="column">The column number where the token starts.</param>
        /// <param name="originalBytes">The original byte sequence for this token. Defaults to an empty array if null.</param>
        public VB6Token(VB6TokenType type, string value, int lineNumber = 0, int column = 0, byte[]? originalBytes = null)
        {
            Type = type;
            Value = value ?? throw new ArgumentNullException(nameof(value));
            LineNumber = lineNumber;
            Column = column;
            OriginalBytes = originalBytes ?? [];
        }

        public override string ToString()
        {
            string bytesStr = OriginalBytes.Length > 0 ? $", Bytes:[{string.Join(" ", OriginalBytes.Select(b => b.ToString("X2")))}]" : "";
            // Escape newlines and tabs in value for better readability
            string displayValue = Value.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
            if (displayValue.Length > 50) displayValue = string.Concat(displayValue.AsSpan(0, 47), "...");

            return $"Token(Type={Type}, Value=\"{displayValue}\", Line={LineNumber}, Col={Column}{bytesStr})";
        }

        public override bool Equals(object? obj)
        {
            if (obj is VB6Token other)
            {
                // For equality, typically compare Type, Value, and potentially position.
                // OriginalBytes might differ due to encoding normalization or if not always captured.
                return Type == other.Type && Value == other.Value && LineNumber == other.LineNumber && Column == other.Column;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Type, Value, LineNumber, Column);
        }

        public static bool operator ==(VB6Token left, VB6Token right) => left.Equals(right);
        public static bool operator !=(VB6Token left, VB6Token right) => !(left == right);
    }
}