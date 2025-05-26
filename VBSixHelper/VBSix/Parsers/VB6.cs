using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using VBSix.Language; // For VB6Token, VB6TokenType
using VBSix.Errors;   // For VB6ErrorKind, VB6ParseException

namespace VBSix.Parsers
{
    public static partial class VB6 // Made partial for potential splitting if Vb6.TokenParsing.cs is used
    {
        /// <summary>
        /// Tries to consume a specific string from the stream.
        /// Throws an exception if the string is not found or the stream is too short.
        /// </summary>
        /// <param name="stream">The input stream.</param>
        /// <param name="toTake">The string to consume.</param>
        /// <param name="caseSensitive">Whether the comparison should be case-sensitive.</param>
        /// <returns>A tuple containing the new stream state and the consumed string.</returns>
        public static (VB6Stream NewStream, string ConsumedString) TakeSpecificString(VB6Stream stream, string toTake, bool caseSensitive = true)
        {
            if (stream.Length < toTake.Length)
            {
                throw stream.CreateException(VB6ErrorKind.KeywordNotFound, // Or a more generic "ExpectedStringNotFound"
                    innerException: new Exception($"Expected \"{toTake}\" but stream is too short (length {stream.Length}). File: {stream.FilePath}, Line: {stream.LineNumber}, Col: {stream.Column}"));
            }

            // Using Encoding.Default for VB6 general string operations as it often aligns with system ANSI.
            // For strict ASCII or specific encodings, adjust accordingly.
            string streamStart = Encoding.Default.GetString(stream.PeekSlice(toTake.Length).Span);
            bool match = caseSensitive ? streamStart.Equals(toTake) : streamStart.Equals(toTake, StringComparison.OrdinalIgnoreCase);

            if (match)
            {
                return (stream.Advance(toTake.Length), toTake);
            }
            
            string foundSnippet = streamStart.Length > 20 ? string.Concat(streamStart.AsSpan(0, 20), "...") : streamStart;
            throw stream.CreateException(VB6ErrorKind.KeywordNotFound, // Or "ExpectedStringNotFound"
                innerException: new Exception($"Expected \"{toTake}\" but found \"{foundSnippet}\". File: {stream.FilePath}, Line: {stream.LineNumber}, Col: {stream.Column}"));
        }

        /// <summary>
        /// Consumes and returns characters from the stream until the specified delimiter byte is encountered.
        /// The delimiter itself is not consumed.
        /// </summary>
        public static (VB6Stream NewStream, string Text) TakeUntil(VB6Stream stream, byte delimiter)
        {
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || b.Value == delimiter)
                {
                    break;
                }
                length++;
            }
            // Using Encoding.Default which usually maps to the system's ANSI codepage (e.g., Windows-1252 for Western languages)
            // This is generally more appropriate for VB6 source code than pure ASCII.
            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            string text = registeredEncoding.GetString(stream.PeekSlice(length).Span);
            return (stream.Advance(length), text);
        }

        /// <summary>
        /// Consumes a single specific character (byte) from the stream.
        /// Throws an exception if the character is not found at the current position.
        /// </summary>
        public static VB6Stream ParseLiteral(VB6Stream stream, byte character)
        {
            var checkpoint = stream.Checkpoint(); // For error reporting with original position
            if (stream.IsEmpty || stream.PeekByte() != character)
            {
                char expectedChar = (char)character;
                string foundCharStr = stream.IsEmpty ? "[EOF]" : $"'{Convert.ToChar(stream.PeekByte()!.Value)}' (0x{stream.PeekByte()!.Value:X2})";
                throw stream.CreateExceptionFromCheckpoint(checkpoint, VB6ErrorKind.Unparseable, // Or a more specific "ExpectedLiteralNotFound"
                    innerException: new Exception($"Expected character '{expectedChar}' (0x{character:X2}) but found {foundCharStr}. File: {stream.FilePath}"));
            }
            return stream.Advance(1);
        }

        /// <summary>
        /// Consumes characters until a line ending (CR/LF) or a comment initiator (apostrophe) is found.
        /// The line ending or comment initiator is not consumed by this method.
        /// </summary>
        public static (VB6Stream NewStream, string Text) TakeUntilLineEndingOrComment(VB6Stream stream)
        {
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || b.Value == (byte)'\r' || b.Value == (byte)'\n' || b.Value == (byte)'\'')
                {
                    break;
                }
                length++;
            }
            string text = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            return (stream.Advance(length), text);
        }
        
        /// <summary>
        /// Consumes characters until a line ending (CR/LF) is found.
        /// The line ending characters are not consumed by this method.
        /// The `nextStream` in the return tuple is the stream advanced past the consumed text.
        /// </summary>
        public static (VB6Stream nextStream, string text) TakeUntilLineEnding(VB6Stream stream, out VB6Stream dummyMatchSignature)
        {
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || b.Value == (byte)'\r' || b.Value == (byte)'\n')
                {
                    break;
                }
                length++;
            }
            string text = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            dummyMatchSignature = stream.Advance(length);
            return (stream.Advance(length), text);
        }


        /// <summary>
        /// Skips characters from the stream as long as the predicate is true.
        /// </summary>
        public static VB6Stream SkipWhile(VB6Stream stream, Func<byte, bool> predicate)
        {
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || !predicate(b.Value))
                {
                    break;
                }
                length++;
            }
            return stream.Advance(length);
        }
        
        /// <summary>
        /// Takes characters from the stream as long as the predicate is true.
        /// </summary>
        public static (VB6Stream NewStream, string Text) TakeWhile(VB6Stream stream, Func<byte, bool> predicate)
        {
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || !predicate(b.Value))
                {
                    break;
                }
                length++;
            }
            string text = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            return (stream.Advance(length), text);
        }
        
        /// <summary>
        /// Parses a VB6 string literal, handling escaped double quotes ("").
        /// </summary>
        public static (VB6Stream NewStream, string Content) ParseVB6String(VB6Stream stream)
        {
            var initialCheckpoint = stream.Checkpoint();
            if (stream.IsEmpty || stream.PeekByte() != (byte)'"')
            {
                throw stream.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.StringParseError, innerException: new Exception("String literal must start with a double quote."));
            }

            VB6Stream s = stream.Advance(1); // Consume opening quote
            
            StringBuilder contentBuilder = new();
            bool unterminated = true;

            while (!s.IsEmpty)
            {
                byte currentByte = s.PeekByte().Value; // IsEmpty check done

                if (currentByte == (byte)'"') // Found a quote
                {
                    s = s.Advance(1); // Consume this quote
                    if (!s.IsEmpty && s.PeekByte() == (byte)'"') // It's an escaped ""
                    {
                        contentBuilder.Append('"'); // Add one quote to content
                        s = s.Advance(1); // Consume the second quote of the "" pair
                    }
                    else // It was a single quote, so it's the closing quote for the string
                    {
                        unterminated = false;
                        break; 
                    }
                }
                else // Not a quote, just a regular character in the string
                {
                    // Handle line breaks within strings if VB6 allows (typically not directly without concatenation)
                    // For now, assuming standard characters. VB uses Windows-1252 or system ANSI.
                    Encoding registeredEncoding = Encoding.GetEncoding(1252);
                    contentBuilder.Append(registeredEncoding.GetString([currentByte]));
                    s = s.Advance(1); // Consume this character
                }
            }

            if (unterminated)
            {
                throw stream.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.UnterminatedString);
            }
            return (s, contentBuilder.ToString());
        }

        public static VB6Token? LineCommentParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty || stream.PeekByte() != (byte)'\'') return null;

            var checkpoint = stream.Checkpoint();
            
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || b.Value == (byte)'\r' || b.Value == (byte)'\n') break;
                length++;
            }
            string text = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            newStream = stream.Advance(length);
            return new VB6Token(VB6TokenType.Comment, text, checkpoint.LineNumber, checkpoint.Column);
        }

        public static VB6Token? RemCommentParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            var checkpoint = stream.Checkpoint();
            if (!stream.CompareSliceCaseless("REM")) return null;

            // Ensure "REM" is a whole word or followed by non-alphanumeric
            if (stream.Length > 3)
            {
                byte? nextCharByte = stream.PeekByteAt(3);
                if (nextCharByte.HasValue && (char.IsLetterOrDigit((char)nextCharByte.Value) || (char)nextCharByte.Value == '_'))
                    return null; // e.g., "REMINDER"
            }
            
            // Consume "REM"
            VB6Stream sAfterRem = stream.Advance(3);
            
            int length = 0;
            while (length < sAfterRem.Length)
            {
                byte? b = sAfterRem.PeekByteAt(length);
                if (!b.HasValue || b.Value == (byte)'\r' || b.Value == (byte)'\n') break;
                length++;
            }
            string commentTextAfterRem = Encoding.Default.GetString(sAfterRem.PeekSlice(length).Span);
            newStream = sAfterRem.Advance(length);
            return new VB6Token(VB6TokenType.Comment, "REM" + commentTextAfterRem, checkpoint.LineNumber, checkpoint.Column);
        }
        
        public static bool IsVbIdentifierStartChar(byte b)
        {
            char c = (char)b;
            // VB6 allows letters (including extended ANSI) as the first character.
            return char.IsLetter(c) || (c >= 128 && c <= 255);
        }

        public static bool IsVbIdentifierPartChar(byte b)
        {
            char c = (char)b;
            // VB6 allows letters, digits, and underscore (also extended ANSI chars) in identifiers.
            return char.IsLetterOrDigit(c) || c == '_' || (c >= 128 && c <= 255);
        }

        // Helper method to parse one or more whitespace characters (space, tab).
        public static VB6Token? WhitespaceParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty) return null;

            var checkpoint = stream.Checkpoint();
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || (b.Value != (byte)' ' && b.Value != (byte)'\t'))
                {
                    break;
                }
                length++;
            }

            if (length == 0) return null;

            string wsText = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            newStream = stream.Advance(length);
            return new VB6Token(VB6TokenType.Whitespace, wsText, checkpoint.LineNumber, checkpoint.Column);
        }

        // Helper method to parse a newline sequence (CRLF, LF, or CR).
        public static VB6Token? NewlineParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            if (stream.IsEmpty) return null;

            var checkpoint = stream.Checkpoint();
            byte? firstByte = stream.PeekByte();

            if (firstByte == (byte)'\r')
            {
                if (stream.Length > 1 && stream.PeekByteAt(1) == (byte)'\n') // CRLF
                {
                    newStream = stream.Advance(2);
                    return new VB6Token(VB6TokenType.Newline, "\r\n", checkpoint.LineNumber, checkpoint.Column);
                }
                else // CR
                {
                    newStream = stream.Advance(1);
                    return new VB6Token(VB6TokenType.Newline, "\r", checkpoint.LineNumber, checkpoint.Column);
                }
            }
            else if (firstByte == (byte)'\n') // LF
            {
                newStream = stream.Advance(1);
                return new VB6Token(VB6TokenType.Newline, "\n", checkpoint.LineNumber, checkpoint.Column);
            }
            return null;
        }

        public static VB6Token? VariableNameParse(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            var checkpoint = stream.Checkpoint();
            if (stream.IsEmpty) return null;
            
            byte? firstByte = stream.PeekByte();
            if (!firstByte.HasValue || !IsVbIdentifierStartChar(firstByte.Value)) return null;

            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || !IsVbIdentifierPartChar(b.Value)) break;
                length++;
            }

            if (length == 0) return null; // Should be caught by firstByte check
            if (length >= 255) // VB6 limit is 255 characters for identifiers
                throw stream.CreateExceptionFromCheckpoint(checkpoint, VB6ErrorKind.VariableNameTooLong);

            string name = Encoding.Default.GetString(stream.PeekSlice(length).Span);
            newStream = stream.Advance(length);
            return new VB6Token(VB6TokenType.Identifier, name, checkpoint.LineNumber, checkpoint.Column);
        }
        
        public static bool IsEnglishCode(ReadOnlyMemory<byte> content)
        {
            if (content.IsEmpty) return true;
            int characterCount = content.Length;
            if (characterCount == 0) return true; // Avoid division by zero for safety

            int higherHalfCharacterCount = 0;
            for(int i=0; i < content.Span.Length; i++)
            {
                if (content.Span[i] >= 128) higherHalfCharacterCount++;
            }
            if (higherHalfCharacterCount == 0) return true;
            
            return (100.0 * higherHalfCharacterCount / characterCount) < 1.0;
        }

        /// <summary>
        /// Attempts to parse the next single VB6 token from the stream.
        /// This is an internal helper for the main Vb6Parse loop.
        /// </summary>
        private static VB6Token? TryParseNextSingleToken(VB6Stream stream, out VB6Stream newStream)
        {
            newStream = stream;
            VB6Token? token;

            // Order of these checks is critical.
            // 1. Whitespace (most frequent, simple check)
            if ((token = WhitespaceParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 2. Newline (also frequent structural element)
            if ((token = NewlineParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 3. Comments (consume until EOL)
            if ((token = LineCommentParse(stream, out newStream)) != null) return token;
            newStream = stream;
            if ((token = RemCommentParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 4. String Literals (can contain almost anything, so parse before keywords/identifiers if ambiguous)
            //    The main Vb6Parse loop handles string literals separately due to their content.
            //    If called from Vb6Parse, string literals would have been handled already.
            //    If this method is to be more general, string literal parsing needs to be here.
            //    For now, assuming strings are handled by the caller (Vb6Parse main loop).
            //    If firstByte is '"', Vb6Parse handles it.

            // 5. Keywords (longer keywords should be checked before shorter prefixes)
            if ((token = VB6Tokens.TryParseKeywordToken(stream, out newStream)) != null) return token;
            newStream = stream;

            // 6. Numbers (can include prefixes like &H, &O and suffixes)
            if ((token = VB6Tokens.NumberParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 7. Identifiers (must come *after* keywords)
            if ((token = VariableNameParse(stream, out newStream)) != null) return token;
            newStream = stream;

            // 8. Symbols and Operators (multi-character ones first)
            if ((token = VB6Tokens.Vb6SymbolParse(stream, out newStream)) != null) return token;
            newStream = stream;

            return null; // No token recognized at the current position
        }

        public static List<VB6Token> Vb6Parse(VB6Stream stream)
        {
            var tokens = new List<VB6Token>();
            VB6Stream currentStream = stream;

            if (!IsEnglishCode(currentStream.CurrentStream)) 
                throw currentStream.CreateException(VB6ErrorKind.LikelyNonEnglishCharacterSet);

            while (!currentStream.IsEmpty)
            {
                var checkpoint = currentStream.Checkpoint();
                byte? firstByte = currentStream.PeekByte();

                if (firstByte == 0) break; // Null byte often signals end of text data in some contexts

                VB6Stream nextStreamAfterToken;
                VB6Token? token = null;

                // Order of attempts is crucial
                if ((token = NewlineParse(currentStream, out nextStreamAfterToken)) != null) { /* Handled */ }
                else if ((token = LineCommentParse(currentStream, out nextStreamAfterToken)) != null) { /* Handled */ }
                else if ((token = RemCommentParse(currentStream, out nextStreamAfterToken)) != null) { /* Handled */ }
                // StringParse needs careful placement. If Vb6TokenParse handles keywords first, "As" in "Dim x As String" is fine.
                // If strings can contain keywords, string parsing might need to be more context-aware or happen earlier.
                else if (firstByte == (byte)'"') // Check for string literal start
                {
                    var stringParseResult = ParseVB6String(currentStream); // This returns content *without* quotes
                    token = new VB6Token(VB6TokenType.LiteralString, "\"" + stringParseResult.Content + "\"", checkpoint.LineNumber, checkpoint.Column); // Add quotes back for the token Value
                    nextStreamAfterToken = stringParseResult.NewStream;
                }
                else if ((token = VB6Tokens.Vb6TokenParse(currentStream, out nextStreamAfterToken)) != null) { /* Handled */ }
                
                if (token != null)
                {
                    tokens.Add(token);
                    currentStream = nextStreamAfterToken;
                }
                else
                {
                    if (!currentStream.IsEmpty) // If no token parsed and stream not empty, it's an error
                        throw currentStream.CreateException(VB6ErrorKind.UnknownToken, innerException: new Exception($"Unrecognized sequence starting with '{(char)currentStream.PeekByte()!}' (0x{currentStream.PeekByte()!:X2})"));
                    break; 
                }
            }
            return tokens;
        }
    }
}