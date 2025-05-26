using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using VBSix.Language;
using VBSix.Errors;

namespace VBSix.Parsers
{
    public static partial class VB6Tokens
    {
        // Keyword list for TryParseKeywordToken
        // Order by length descending for multi-word keywords or prefix issues.
        private static readonly List<Tuple<string, VB6TokenType>> KnownKeywords = [.. new List<Tuple<string, VB6TokenType>>()
        {
            // Longer keywords first
            Tuple.Create("AddressOf", VB6TokenType.AddressOfKeyword), Tuple.Create("AppActivate", VB6TokenType.AppActivateKeyword),
            Tuple.Create("DefBool", VB6TokenType.DefBoolKeyword), Tuple.Create("DefByte", VB6TokenType.DefByteKeyword),
            Tuple.Create("DefCur", VB6TokenType.DefCurKeyword), Tuple.Create("DefDate", VB6TokenType.DefDateKeyword),
            Tuple.Create("DefDbl", VB6TokenType.DefDblKeyword), Tuple.Create("DefDec", VB6TokenType.DefDecKeyword),
            Tuple.Create("DefInt", VB6TokenType.DefIntKeyword), Tuple.Create("DefLng", VB6TokenType.DefLngKeyword),
            Tuple.Create("DefObj", VB6TokenType.DefObjKeyword), Tuple.Create("DefSng", VB6TokenType.DefSngKeyword),
            Tuple.Create("DefStr", VB6TokenType.DefStrKeyword), Tuple.Create("DefVar", VB6TokenType.DefVarKeyword),
            Tuple.Create("DeleteSetting", VB6TokenType.DeleteSettingKeyword), Tuple.Create("ElseIf", VB6TokenType.ElseIfKeyword),
            Tuple.Create("FileCopy", VB6TokenType.FileCopyKeyword), Tuple.Create("Implements", VB6TokenType.ImplementsKeyword),
            Tuple.Create("Line Input", VB6TokenType.InputKeyword), // Special case, might need combined parsing
            Tuple.Create("ParamArray", VB6TokenType.ParamArrayKeyword), Tuple.Create("Property Get", VB6TokenType.PropertyKeyword),
            Tuple.Create("Property Let", VB6TokenType.PropertyKeyword), Tuple.Create("Property Set", VB6TokenType.PropertyKeyword),
            Tuple.Create("RaiseEvent", VB6TokenType.RaiseEventKeyword), Tuple.Create("SavePicture", VB6TokenType.SavePictureKeyword),
            Tuple.Create("SaveSetting", VB6TokenType.SaveSettingKeyword), Tuple.Create("SendKeys", VB6TokenType.SendKeysKeyword),
            Tuple.Create("SetAttr", VB6TokenType.SetAttrKeyword), Tuple.Create("WithEvents", VB6TokenType.WithEventsKeyword),

            // Single word keywords
            Tuple.Create("Alias", VB6TokenType.AliasKeyword), Tuple.Create("And", VB6TokenType.AndKeyword),
            Tuple.Create("As", VB6TokenType.AsKeyword), Tuple.Create("Base", VB6TokenType.BaseKeyword),
            Tuple.Create("Beep", VB6TokenType.BeepKeyword), Tuple.Create("Binary", VB6TokenType.BinaryKeyword),
            Tuple.Create("Boolean", VB6TokenType.BooleanKeyword), Tuple.Create("ByRef", VB6TokenType.ByRefKeyword),
            Tuple.Create("ByVal", VB6TokenType.ByValKeyword), Tuple.Create("Byte", VB6TokenType.ByteKeyword),
            Tuple.Create("Call", VB6TokenType.CallKeyword), Tuple.Create("Case", VB6TokenType.CaseKeyword),
            Tuple.Create("ChDir", VB6TokenType.ChDirKeyword), Tuple.Create("ChDrive", VB6TokenType.ChDriveKeyword),
            Tuple.Create("Close", VB6TokenType.CloseKeyword), Tuple.Create("Compare", VB6TokenType.CompareKeyword),
            Tuple.Create("Const", VB6TokenType.ConstKeyword), Tuple.Create("Currency", VB6TokenType.CurrencyKeyword),
            Tuple.Create("Date", VB6TokenType.DateKeyword), Tuple.Create("Decimal", VB6TokenType.DecimalKeyword),
            Tuple.Create("Declare", VB6TokenType.DeclareKeyword), Tuple.Create("Dim", VB6TokenType.DimKeyword),
            Tuple.Create("Do", VB6TokenType.DoKeyword), Tuple.Create("Double", VB6TokenType.DoubleKeyword),
            Tuple.Create("Else", VB6TokenType.ElseKeyword), Tuple.Create("Empty", VB6TokenType.EmptyKeyword),
            Tuple.Create("End", VB6TokenType.EndKeyword), Tuple.Create("Enum", VB6TokenType.EnumKeyword),
            Tuple.Create("Eqv", VB6TokenType.EqvKeyword), Tuple.Create("Erase", VB6TokenType.EraseKeyword),
            Tuple.Create("Error", VB6TokenType.ErrorKeyword), Tuple.Create("Event", VB6TokenType.EventKeyword),
            Tuple.Create("Exit", VB6TokenType.ExitKeyword), Tuple.Create("Explicit", VB6TokenType.ExplicitKeyword),
            Tuple.Create("False", VB6TokenType.FalseKeyword), Tuple.Create("For", VB6TokenType.ForKeyword),
            Tuple.Create("Friend", VB6TokenType.FriendKeyword), Tuple.Create("Function", VB6TokenType.FunctionKeyword),
            Tuple.Create("Get", VB6TokenType.GetKeyword), Tuple.Create("Goto", VB6TokenType.GotoKeyword),
            Tuple.Create("If", VB6TokenType.IfKeyword), Tuple.Create("Imp", VB6TokenType.ImpKeyword),
            Tuple.Create("Input", VB6TokenType.InputKeyword), Tuple.Create("Integer", VB6TokenType.IntegerKeyword),
            Tuple.Create("Is", VB6TokenType.IsKeyword), Tuple.Create("Kill", VB6TokenType.KillKeyword),
            Tuple.Create("Len", VB6TokenType.LenKeyword), Tuple.Create("Let", VB6TokenType.LetKeyword),
            Tuple.Create("Lib", VB6TokenType.LibKeyword), Tuple.Create("Like", VB6TokenType.LikeKeyword),
            Tuple.Create("Line", VB6TokenType.LineKeyword), Tuple.Create("Load", VB6TokenType.LoadKeyword),
            Tuple.Create("Lock", VB6TokenType.LockKeyword), Tuple.Create("Long", VB6TokenType.LongKeyword),
            Tuple.Create("LSet", VB6TokenType.LSetKeyword), Tuple.Create("Me", VB6TokenType.MeKeyword),
            Tuple.Create("Mid", VB6TokenType.MidKeyword), Tuple.Create("MkDir", VB6TokenType.MkDirKeyword),
            Tuple.Create("Mod", VB6TokenType.ModKeyword), Tuple.Create("Name", VB6TokenType.NameKeyword),
            Tuple.Create("New", VB6TokenType.NewKeyword), Tuple.Create("Next", VB6TokenType.NextKeyword),
            Tuple.Create("Not", VB6TokenType.NotKeyword), Tuple.Create("Null", VB6TokenType.NullKeyword),
            Tuple.Create("Object", VB6TokenType.ObjectKeyword), Tuple.Create("On", VB6TokenType.OnKeyword),
            Tuple.Create("Open", VB6TokenType.OpenKeyword), Tuple.Create("Option", VB6TokenType.OptionKeyword),
            Tuple.Create("Optional", VB6TokenType.OptionalKeyword), Tuple.Create("Or", VB6TokenType.OrKeyword),
            Tuple.Create("Preserve", VB6TokenType.PreserveKeyword), Tuple.Create("Print", VB6TokenType.PrintKeyword),
            Tuple.Create("Private", VB6TokenType.PrivateKeyword), Tuple.Create("Property", VB6TokenType.PropertyKeyword),
            Tuple.Create("Public", VB6TokenType.PublicKeyword), Tuple.Create("Put", VB6TokenType.PutKeyword),
            Tuple.Create("Randomize", VB6TokenType.RandomizeKeyword), Tuple.Create("ReDim", VB6TokenType.ReDimKeyword),
            Tuple.Create("Reset", VB6TokenType.ResetKeyword), Tuple.Create("Resume", VB6TokenType.ResumeKeyword),
            Tuple.Create("RmDir", VB6TokenType.RmDirKeyword), Tuple.Create("RSet", VB6TokenType.RSetKeyword),
            Tuple.Create("Seek", VB6TokenType.SeekKeyword), Tuple.Create("Select", VB6TokenType.SelectKeyword),
            Tuple.Create("Set", VB6TokenType.SetKeyword), Tuple.Create("Single", VB6TokenType.SingleKeyword),
            Tuple.Create("Static", VB6TokenType.StaticKeyword), Tuple.Create("Step", VB6TokenType.StepKeyword),
            Tuple.Create("Stop", VB6TokenType.StopKeyword), Tuple.Create("String", VB6TokenType.StringKeyword),
            Tuple.Create("Sub", VB6TokenType.SubKeyword), Tuple.Create("Then", VB6TokenType.ThenKeyword),
            Tuple.Create("Time", VB6TokenType.TimeKeyword), Tuple.Create("To", VB6TokenType.ToKeyword),
            Tuple.Create("True", VB6TokenType.TrueKeyword), Tuple.Create("Type", VB6TokenType.TypeKeyword),
            Tuple.Create("Unlock", VB6TokenType.UnlockKeyword), Tuple.Create("Until", VB6TokenType.UntilKeyword),
            Tuple.Create("Variant", VB6TokenType.VariantKeyword), Tuple.Create("Wend", VB6TokenType.WendKeyword),
            Tuple.Create("While", VB6TokenType.WhileKeyword), Tuple.Create("Width", VB6TokenType.WidthKeyword),
            Tuple.Create("With", VB6TokenType.WithKeyword), Tuple.Create("Write", VB6TokenType.WriteKeyword),
            Tuple.Create("Xor", VB6TokenType.XorKeyword)
        }.OrderByDescending(t => t.Item1.Length)]; // Sort keywords by length descending to handle cases like "ElseIf" before "Else" correctly.


        public static VB6Token? TryParseKeywordToken(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            var checkpoint = stream.Checkpoint();

            foreach (var entry in KnownKeywords)
            {
                string keyword = entry.Item1;
                VB6TokenType type = entry.Item2;

                if (stream.Length >= keyword.Length && stream.CompareSliceCaseless(keyword))
                {
                    // Boundary check: Ensure it's not a prefix of a larger identifier
                    if (stream.Length > keyword.Length)
                    {
                        byte? nextByte = stream.PeekByteAt(keyword.Length);
                        if (nextByte.HasValue && VB6.IsVbIdentifierPartChar(nextByte.Value))
                        {
                            continue; // It's a prefix, not the keyword itself
                        }
                    }

                    // It's a match
                    string actualValue = Encoding.Default.GetString(stream.PeekSlice(keyword.Length).Span);
                    newStream = stream.Advance(keyword.Length);
                    return new VB6Token(type, actualValue, checkpoint.LineNumber, checkpoint.Column);
                }
            }
            return null; // No keyword matched
        }

        public static VB6Token? Vb6SymbolParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty) return null;
            var checkpoint = stream.Checkpoint();

            // Try multi-character symbols first
            if (stream.Length >= 2)
            {
                string twoChar = Encoding.Default.GetString(stream.PeekSlice(2).Span);
                VB6TokenType? type = twoChar switch
                {
                    "<>" => VB6TokenType.NotEqualOperator,
                    "<=" => VB6TokenType.LessThanOrEqualOperator,
                    ">=" => VB6TokenType.GreaterThanOrEqualOperator,
                    ":=" => VB6TokenType.AssignmentOperator, // While VB uses '=', if := was supported
                    // Add other multi-character VB6 operators if any (e.g., // for comments in some contexts, but handled elsewhere)
                    _ => null
                };
                if (type.HasValue)
                {
                    newStream = stream.Advance(2);
                    return new VB6Token(type.Value, twoChar, checkpoint.LineNumber, checkpoint.Column);
                }
            }

            // Single character symbols
            char singleChar = (char)stream.PeekByte().Value; // IsEmpty check done

            if (singleChar == '_')
            {
                // Special case for line continuation character '_'
                // It can be followed by whitespace or newline, so we handle it separately
                if (stream.Length > 1 && (stream.PeekByteAt(1) == (byte)' ' || stream.PeekByteAt(1) == (byte)'\t'))
                {
                    newStream = stream.Advance(2); // Consume '_' and whitespace
                    return new VB6Token(VB6TokenType.LineContinuation, "_", checkpoint.LineNumber, checkpoint.Column);
                }
                else if (stream.Length > 1 && (stream.PeekByteAt(1) == (byte)'\r' || stream.PeekByteAt(1) == (byte)'\n'))
                {
                    newStream = stream.Advance(2); // Consume '_' and newline
                    return new VB6Token(VB6TokenType.LineContinuation, "_", checkpoint.LineNumber, checkpoint.Column);
                }
            }

            VB6TokenType? singleType = singleChar switch
            {
                '=' => VB6TokenType.EqualityOperator, '<' => VB6TokenType.LessThanOperator,
                '>' => VB6TokenType.GreaterThanOperator, '(' => VB6TokenType.LeftParenthesis,
                ')' => VB6TokenType.RightParenthesis, '[' => VB6TokenType.LeftSquareBracket,
                ']' => VB6TokenType.RightSquareBracket, ',' => VB6TokenType.Comma,
                ';' => VB6TokenType.Semicolon, ':' => VB6TokenType.Colon,
                '+' => VB6TokenType.PlusOperator, '-' => VB6TokenType.MinusOperator,
                '*' => VB6TokenType.MultiplyOperator, '/' => VB6TokenType.DivideOperator,
                '\\' => VB6TokenType.IntegerDivideOperator, '^' => VB6TokenType.ExponentiateOperator,
                '&' => VB6TokenType.Ampersand, '.' => VB6TokenType.Dot,
                '!' => VB6TokenType.ExclamationMark, '#' => VB6TokenType.Octothorpe,
                '$' => VB6TokenType.DollarSign, '%' => VB6TokenType.Percent,
                '@' => VB6TokenType.AtSign, '_' => VB6TokenType.Underscore,
                _ => null
            };

            if (singleType.HasValue)
            {
                newStream = stream.Advance(1);
                return new VB6Token(singleType.Value, singleChar.ToString(), checkpoint.LineNumber, checkpoint.Column);
            }
            return null;
        }
        
        public static VB6Token? WhitespaceParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            var checkpoint = stream.Checkpoint();
            int length = 0;
            while(length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if(b.HasValue && (b.Value == (byte)' ' || b.Value == (byte)'\t'))
                    length++;
                else
                    break;
            }
            if (length > 0)
            {
                string ws = Encoding.Default.GetString(stream.PeekSlice(length).Span);
                newStream = stream.Advance(length);
                return new VB6Token(VB6TokenType.Whitespace, ws, checkpoint.LineNumber, checkpoint.Column);
            }
            return null;
        }
        
        public static VB6Token? NewlineParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty) return null;
            var checkpoint = stream.Checkpoint();

            byte? first = stream.PeekByte();
            if (first == (byte)'\r')
            {
                if (stream.Length > 1 && stream.PeekByteAt(1) == (byte)'\n')
                {
                    newStream = stream.Advance(2);
                    return new VB6Token(VB6TokenType.Newline, "\r\n", checkpoint.LineNumber, checkpoint.Column);
                }
                newStream = stream.Advance(1);
                return new VB6Token(VB6TokenType.Newline, "\r", checkpoint.LineNumber, checkpoint.Column);
            }
            if (first == (byte)'\n')
            {
                newStream = stream.Advance(1);
                return new VB6Token(VB6TokenType.Newline, "\n", checkpoint.LineNumber, checkpoint.Column);
            }
            return null;
        }

        public static VB6Token? NumberParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty) return null;
            var checkpoint = stream.Checkpoint();

            int len = 0;
            bool isHexOrOct = false;
            bool hasDecimalPoint = false;
            bool hasExponent = false;

            // Check for &H or &O prefix
            if (stream.Length >= 2 && stream.PeekByteAt(0) == (byte)'&')
            {
                byte? second = stream.PeekByteAt(1);
                if (second == (byte)'H' || second == (byte)'h') { isHexOrOct = true; len = 2; }
                else if (second == (byte)'O' || second == (byte)'o') { isHexOrOct = true; len = 2; }
            }

            // Parse digits
            while (len < stream.Length)
            {
                byte b = stream.PeekByteAt(len).Value;
                if (isHexOrOct)
                {
                    if (!((b >= (byte)'0' && b <= (byte)'9') ||
                          (b >= (byte)'A' && b <= (byte)'F') ||
                          (b >= (byte)'a' && b <= (byte)'f') ||
                          ( (stream.PeekByteAt(1) == 'O' || stream.PeekByteAt(1) == 'o') && b >= (byte)'0' && b <= (byte)'7' ) )) // Octal digits for &O
                        break;
                }
                else // Decimal number
                {
                    if (b >= (byte)'0' && b <= (byte)'9') { /* digit */ }
                    else if (b == (byte)'.' && !hasDecimalPoint && !hasExponent) { hasDecimalPoint = true; }
                    else if ((b == (byte)'E' || b == (byte)'e') && !hasExponent)
                    {
                        hasExponent = true;
                        // Optional sign for exponent
                        if (len + 1 < stream.Length && (stream.PeekByteAt(len + 1) == (byte)'+' || stream.PeekByteAt(len + 1) == (byte)'-'))
                            len++; // Consume sign
                    }
                    else break;
                }
                len++;
            }

            // Type suffixes for decimal numbers
            if (!isHexOrOct && !hasExponent && len < stream.Length)
            {
                byte suffix = stream.PeekByteAt(len).Value;
                if ("%&!#@".Contains((char)suffix))
                    len++;
            }

            if (len == 0 || (isHexOrOct && len <= 2)) return null; // No actual number digits

            string numStr = Encoding.Default.GetString(stream.PeekSlice(len).Span);
            // Additional validation to prevent parsing just "." or "E" as numbers
            if (!isHexOrOct && numStr == "." ) return null;
            if (!isHexOrOct && numStr.EndsWith("E", StringComparison.OrdinalIgnoreCase) && !char.IsDigit(numStr[^2])) return null;


            newStream = stream.Advance(len);
            return new VB6Token(VB6TokenType.LiteralNumber, numStr, checkpoint.LineNumber, checkpoint.Column);
        }

        public static VB6Token? Vb6TokenParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream; // Default to no change

            // The order of these checks is crucial to avoid mis-tokenization.
            // E.g., Keywords before Identifiers, specific symbols before general.
            VB6Token? token;

            // 1. Whitespace (must be handled to separate other tokens)
            if ((token = WhitespaceParse(stream, out newStream)) != null) return token;
            newStream = stream; 
            
            // 2. Comments (Line and REM) - These consume till EOL
            // Already handled by Vb6Parse outer loop, typically. But can be here for completeness.
            // if ((token = LineCommentParse(stream, out newStream)) != null) return token;
            // newStream = stream;
            // if ((token = RemCommentParse(stream, out newStream)) != null) return token;
            // newStream = stream;

            // 3. Keywords
            if ((token = TryParseKeywordToken(stream, out newStream)) != null) return token;
            newStream = stream;
            
            // 4. Numbers
            if ((token = NumberParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 5. Identifiers (must be after keywords)
            if ((token = VB6.VariableNameParse(stream, out newStream)) != null) return token;
            newStream = stream;
            
            // 6. Symbols and Operators
            if ((token = Vb6SymbolParse(stream, out newStream)) != null) return token;
            newStream = stream;
            
            // 7. String Literals are handled by the main Vb6Parse loop because they can contain anything.
            // If Vb6Parse doesn't pre-handle them, it's complex here.
            
            // If nothing else matched, it's an unknown token or end of meaningful stream.
            return null;
        }
    }
}