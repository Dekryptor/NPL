namespace VB6Parse.Language
{
    /// <summary>
    /// Defines the types of tokens in the VB6 language.
    /// </summary>
    public enum Vb6TokenType
    {
        // --- Special Tokens ---
        Unknown,          // For unrecognized tokens
        EndOfFile,        // To mark the end of input

        // --- Whitespace and Comments ---
        Whitespace,       // Represents one or more whitespace characters
        Newline,          // Represents a newline sequence (CR, LF, CRLF)
        Comment,          // Represents a single-quote comment (e.g., ' This is a comment)
        RemComment,       // Represents a REM statement comment (e.g., REM This is also a comment)

        // --- Literals ---
        StringLiteral,    // e.g., "Hello, World!"
        NumberLiteral,    // e.g., 123, 3.14, &HFF (will need further parsing to determine exact numeric type)
        DateLiteral,      // e.g., #12/31/2023# (though VB6 often parses these from strings too)

        // --- Identifiers ---
        Identifier,       // e.g., MyVariable, MyFunction

        // --- Keywords (Declarations, Modifiers, Types) ---
        ReDimKeyword,
        PreserveKeyword,
        DimKeyword,
        DeclareKeyword,
        AliasKeyword,
        LibKeyword,
        WithEventsKeyword,
        BaseKeyword,        // Option Base
        CompareKeyword,     // Option Compare
        OptionKeyword,
        ExplicitKeyword,    // Option Explicit
        PrivateKeyword,
        PublicKeyword,
        ConstKeyword,
        AsKeyword,
        ByValKeyword,
        ByRefKeyword,
        OptionalKeyword,
        FunctionKeyword,
        StaticKeyword,
        SubKeyword,
        EndKeyword,         // Used in End Sub, End Function, End Type, End If, etc.
        TrueKeyword,
        FalseKeyword,
        EnumKeyword,
        TypeKeyword,        // For User-Defined Types (UDTs)
        FriendKeyword,
        ImplementsKeyword,
        EventKeyword,
        NewKeyword,
        MeKeyword,
        EmptyKeyword,       // Variant subtype
        NullKeyword,        // Variant subtype
        ParamArrayKeyword,

        // --- Built-in Type Keywords ---
        BooleanKeyword,
        DoubleKeyword,
        CurrencyKeyword,
        DecimalKeyword,
        DateKeyword,        // Data type
        ObjectKeyword,
        VariantKeyword,
        ByteKeyword,
        LongKeyword,
        SingleKeyword,
        StringKeyword,
        IntegerKeyword,

        // --- Control Flow & Conditional Keywords ---
        IfKeyword,
        ElseKeyword,
        ElseIfKeyword,
        ThenKeyword,
        GotoKeyword,
        ExitKeyword,        // Exit Sub, Exit Function, Exit For, Exit Do
        ForKeyword,
        ToKeyword,
        StepKeyword,
        StopKeyword,
        WhileKeyword,
        WendKeyword,
        SelectKeyword,
        CaseKeyword,
        DoKeyword,
        UntilKeyword,
        NextKeyword,        // For...Next, Resume Next

        // --- Operators (Logical, Comparison, Arithmetic, etc.) ---
        AndKeyword,         // Logical/Bitwise And
        OrKeyword,          // Logical/Bitwise Or
        XorKeyword,         // Logical/Bitwise Xor
        ModKeyword,         // Modulo operator
        EqvKeyword,         // Logical Equivalence
        AddressOfKeyword,
        ImpKeyword,         // Logical Implication
        IsKeyword,          // Object comparison
        LikeKeyword,        // String pattern matching
        NotKeyword,         // Logical/Bitwise Not

        // --- File I/O & System Keywords ---
        LockKeyword,
        UnlockKeyword,
        WidthKeyword,       // Width #
        WriteKeyword,       // Write #
        TimeKeyword,        // Time statement/function
        SetAttrKeyword,
        SetKeyword,         // For object assignment
        SendKeysKeyword,
        SeekKeyword,        // Seek statement/function
        SaveSettingKeyword,
        SavePictureKeyword,
        RSetKeyword,
        RmDirKeyword,
        ResumeKeyword,
        ResetKeyword,
        RandomizeKeyword,
        RaiseEventKeyword,
        PutKeyword,         // Put #
        PropertyKeyword,    // Property Get/Let/Set
        PrintKeyword,       // Print #
        OpenKeyword,
        OnKeyword,          // On Error, On...GoTo, On...GoSub
        NameKeyword,        // Name file As
        MkDirKeyword,
        MidKeyword,         // Mid statement (also a function)
        LSetKeyword,
        LoadKeyword,        // Load form/control
        LineKeyword,        // Line Input # (also for graphics)
        InputKeyword,       // Input # (also InputBox function)
        LetKeyword,         // Optional assignment keyword
        KillKeyword,
        GetKeyword,         // Get # (also Property Get)
        FileCopyKeyword,
        ErrorKeyword,       // Error statement (also Err object)
        EraseKeyword,
        DeleteSettingKeyword,
        CloseKeyword,       // Close #
        ChDriveKeyword,
        ChDirKeyword,
        CallKeyword,
        BeepKeyword,
        AppActivateKeyword,
        BinaryKeyword,      // Option Compare Binary, Open ... For Binary
        LenKeyword,         // Len function (also used in Open statement)

        // --- DefType Keywords ---
        DefBoolKeyword,
        DefByteKeyword,
        DefIntKeyword,      // DefInt A-Z
        DefLngKeyword,
        DefCurKeyword,
        DefSngKeyword,
        DefDblKeyword,
        DefDecKeyword,
        DefDateKeyword,
        DefStrKeyword,
        DefObjKeyword,
        DefVarKeyword,

        // --- Punctuation & Operators ---
        DollarSign,         // $ (string type character)
        Underscore,         // _ (line continuation)
        Ampersand,          // & (string concatenation, long type character)
        PercentSign,        // % (integer type character)
        Octothorpe,         // # (date literal delimiter, double type character, file number prefix)
        LeftParenthesis,    // (
        RightParenthesis,   // )
        LeftSquareBracket,  // [ (used with some object libraries or for future use)
        RightSquareBracket, // ]
        Comma,              // ,
        Semicolon,          // ;
        AtSign,             // @ (currency type character)
        ExclamationMark,    // ! (single type character, dictionary access)
        EqualityOperator,   // = (also assignment)
        LessThanOperator,   // <
        GreaterThanOperator,// >
        LessThanOrEqualOperator,    // <= (Composed by lexer from < and =)
        GreaterThanOrEqualOperator, // >= (Composed by lexer from > and =)
        NotEqualOperator,           // <> (Composed by lexer from < and >)
        MultiplicationOperator,     // * (also string length in Dim)
        SubtractionOperator,        // - (also negation)
        AdditionOperator,           // +
        DivisionOperator,           // / (float division)
        BackwardSlashOperator,      // \\ (integer division)
        PeriodOperator,             // . (object member access)
        ColonOperator,              // : (statement separator)
        ExponentiationOperator      // ^
    }
}