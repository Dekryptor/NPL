using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Linq;
using VBSix.Errors;
using VBSix.Language; // For VB6Color (if any property parsing happened here, though unlikely for general headers)

namespace VBSix.Parsers
{
    public static class HeaderParser
    {
        // --- Utility Methods (some might be duplicates from Vb6.cs or could be moved to a shared utility) ---

        public static VB6Stream SkipSpace0(VB6Stream stream)
        {
            return VB6.SkipWhile(stream, b => b == (byte)' ' || b == (byte)'\t');
        }

        public static VB6Stream SkipSpace1(VB6Stream stream)
        {
            var initial = stream.Checkpoint();
            VB6Stream newStream = SkipSpace0(stream);
            if (newStream.OffsetFrom(initial) == 0 && !stream.IsEmpty)
            {
                throw stream.CreateExceptionFromCheckpoint(initial, VB6ErrorKind.InternalParseError, innerException: new Exception("Expected at least one whitespace character."));
            }
            return newStream;
        }

        public static (VB6Stream NewStream, string Value) TakeUntil(VB6Stream stream, byte delimiter)
        {
            return VB6.TakeUntil(stream, delimiter);
        }

        public static (VB6Stream NewStream, string Value) TakeUntilAny(VB6Stream stream, params byte[] delimiters)
        {
            // Reusing Vb6.TakeUntilAny if it exists and is suitable, or implement here:
            int length = 0;
            while (length < stream.Length)
            {
                byte? b = stream.PeekByteAt(length);
                if (!b.HasValue || delimiters.Contains(b.Value)) break;
                length++;
            }
            string value = Encoding.Default.GetString(stream.PeekSlice(length).Span); // Use Default for VB6 general strings
            return (stream.Advance(length), value);
        }

        public static (VB6Stream NewStream, string Value) TakeUntilLineEndingOrComment(VB6Stream stream)
        {
            return VB6.TakeUntilLineEndingOrComment(stream);
        }

        public static (VB6Stream NewStream, string Value) TakeUntilLineEnding(VB6Stream stream)
        {
             return VB6.TakeUntilLineEnding(stream, out _); // Adapting to existing Vb6 method signature
        }

        public static VB6Stream ParseLineEnding(VB6Stream stream)
        {
            // Reusing Vb6.NewlineParse but only advancing the stream, not creating a token.
            var checkpoint = stream.Checkpoint();
            if (VB6.NewlineParse(stream, out VB6Stream newStream) != null)
            {
                return newStream;
            }
            throw stream.CreateExceptionFromCheckpoint(checkpoint, VB6ErrorKind.NoLineEnding);
        }

        public static VB6Stream ParseOptionalLineEnding(VB6Stream stream)
        {
            if (VB6.NewlineParse(stream, out VB6Stream newStream) != null)
            {
                return newStream;
            }
            return stream; // No line ending found, return original stream
        }

        public static VB6Stream ParseOptionalTrailingCommentAndLineEnding(VB6Stream stream)
        {
            VB6Stream s = SkipSpace0(stream);
            if (!s.IsEmpty && s.PeekByte() == (byte)'\'')
            {
                (s, _) = TakeUntilLineEnding(s); // Consumes comment text
            }
            s = ParseOptionalLineEnding(s); // Consumes EOL
            return s;
        }


        // --- Core Header Parsing Logic ---

        /// <summary>
        /// Parses a "VERSION X.Y [KIND]" line from a VB6 file header.
        /// </summary>
        public static (VB6Stream NewStream, VB6FileFormatVersion Version, HeaderKind ActualKind) ParseVersionLine(
            VB6Stream stream,
            HeaderKind expectedKindForContext // Used to guide parsing if no explicit KIND keyword is present (e.g., Forms)
        )
        {
            var initialStreamState = stream.Checkpoint();
            VB6Stream s = SkipSpace0(stream);

            (s, _) = VB6.TakeSpecificString(s, "VERSION", caseSensitive: true); // Throws if not found
            s = SkipSpace1(s);

            var (sAfterMajor, majorStr) = TakeUntil(s, (byte)'.');
            if (!byte.TryParse(majorStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out byte majorVer))
            {
                throw s.CreateExceptionFromCheckpoint(initialStreamState, VB6ErrorKind.MajorVersionUnparseable, innerException: new FormatException($"Invalid major version: {majorStr}"));
            }
            s = VB6.ParseLiteral(sAfterMajor, (byte)'.'); // Consume '.'

            // Minor version can be one or more digits, followed by space or EOL
            var (sAfterMinor, minorStr) = VB6.TakeWhile(s, b => b >= (byte)'0' && b <= (byte)'9');
            if (string.IsNullOrEmpty(minorStr) || !byte.TryParse(minorStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out byte minorVer))
            {
                throw s.CreateExceptionFromCheckpoint(initialStreamState, VB6ErrorKind.MinorVersionUnparseable, innerException: new FormatException($"Invalid minor version: {minorStr}"));
            }
            s = sAfterMinor;

            var version = new VB6FileFormatVersion(majorVer, minorVer);
            HeaderKind actualKind = expectedKindForContext; // Default to context if no keyword follows

            // Check for optional KIND keyword (CLASS, MODULE)
            VB6Stream tempStreamForKeywordCheck = SkipSpace0(s); // Look ahead without consuming crucial space for non-keyword lines
            if (!tempStreamForKeywordCheck.IsEmpty &&
                tempStreamForKeywordCheck.PeekByte() != (byte)'\r' &&
                tempStreamForKeywordCheck.PeekByte() != (byte)'\n' &&
                tempStreamForKeywordCheck.PeekByte() != (byte)'\'')
            {
                s = SkipSpace1(s); // Expect at least one space before the KIND keyword
                if (s.CompareSliceCaseless("CLASS"))
                {
                    (s, _) = VB6.TakeSpecificString(s, "CLASS", caseSensitive: true);
                    actualKind = HeaderKind.Class;
                }
                else if (s.CompareSliceCaseless("MODULE")) // Though BAS files usually don't have VERSION line
                {
                    (s, _) = VB6.TakeSpecificString(s, "MODULE", caseSensitive: true);
                    actualKind = HeaderKind.Module;
                }
                // Forms (and UserControls, etc.) do NOT have a "FORM" keyword on the version line.
                // If expectedKindForContext was Form, and we find other text, it might be an error
                // or a different file type than expected by the caller.
                else if (expectedKindForContext != HeaderKind.Form) // If it's not a form, and we have trailing text that isn't CLASS/MODULE
                {
                     var (temp_s, unknownText) = TakeUntilLineEndingOrComment(s);
                     throw s.CreateExceptionFromCheckpoint(initialStreamState, VB6ErrorKind.KeywordNotFound,
                        innerException: new Exception($"Expected header kind keyword (CLASS, MODULE) for '{expectedKindForContext}', but found '{unknownText.Trim()}'. Form headers should not have extra keywords here."));
                }
            }
            // If it's a Form, UserControl, etc., no keyword is expected after version numbers.
            
            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, version, actualKind);
        }


        /// <summary>
        /// Parses a single "Attribute Key = Value" line.
        /// </summary>
        private static (VB6Stream NewStream, VB6Attribute ParsedAttribute) ParseSingleAttributeLine(VB6Stream stream)
        {
            var initialLineCheckpoint = stream.Checkpoint();
            VB6Stream s = SkipSpace0(stream);

            var (sAfterAttrKeyword, _) = VB6.TakeSpecificString(s, "Attribute", caseSensitive: true); // Throws if "Attribute" not found
            s = SkipSpace1(sAfterAttrKeyword);

            // Attribute Key (e.g., VB_Name, VB_GlobalNameSpace, VB_Ext_KEY)
            var (sAfterKey, keyStr) = TakeUntilAny(s, (byte)' ', (byte)'\t', (byte)'=');
            string attributeKey = keyStr.Trim();
            if (string.IsNullOrEmpty(attributeKey))
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.UnknownAttribute, innerException: new Exception("Attribute key is missing."));
            
            s = SkipSpace0(sAfterKey);
            if (s.IsEmpty || s.PeekByte() != (byte)'=')
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.NoEqualSplit, innerException: new Exception($"Missing '=' after attribute key '{attributeKey}'."));
            s = SkipSpace0(s.Advance(1)); // Consume '=' and spaces

            string attributeValue;
            if (attributeKey.Equals("VB_Ext_KEY", StringComparison.OrdinalIgnoreCase))
            {
                // VB_Ext_KEY = "KeyString", "ValueString"
                var (sAfterKeyStr, keyString) = VB6.ParseVB6String(s); // Assuming Vb6.cs has ParseVB6String
                s = SkipSpace0(sAfterKeyStr);
                if (s.IsEmpty || s.PeekByte() != (byte)',')
                    throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.KeyValueParseError, innerException: new Exception("Missing comma in VB_Ext_KEY attribute."));
                s = SkipSpace0(s.Advance(1));
                var (sAfterValStr, valueString) = VB6.ParseVB6String(s);
                // For VB_Ext_KEY, the "Value" of our VB6Attribute will be composite or handled specially.
                attributeValue = $"\"{keyString}\", \"{valueString}\""; // Store raw representation for now
                s = sAfterValStr;
            }
            else if (!s.IsEmpty && s.PeekByte() == (byte)'"') // Quoted string value
            {
                var (sAfterString, parsedString) = VB6.ParseVB6String(s);
                attributeValue = parsedString;
                s = sAfterString;
            }
            else // Unquoted value (e.g., True, False, 0, -1)
            {
                var (sAfterValue, val) = TakeUntilLineEndingOrComment(s);
                attributeValue = val.Trim();
                s = sAfterValue;
            }
            
            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, new VB6Attribute(attributeKey, attributeValue));
        }

        /// <summary>
        /// Parses all "Attribute" lines in a header section.
        /// </summary>
        public static (VB6Stream NewStream, VB6FileAttributes Attributes) ParseAttributes(VB6Stream stream)
        {
            var attributes = new VB6FileAttributes();
            VB6Stream currentStream = stream;
            bool nameFound = false;
            bool attributesEncountered = false; // To track if any Attribute line was processed

            while(!currentStream.IsEmpty)
            {
                var lineStartCheckpoint = currentStream.Checkpoint();
                VB6Stream tempStream = SkipSpace0(currentStream);
                
                if (tempStream.IsEmpty || !tempStream.CompareSliceCaseless("Attribute"))
                {
                    break; // No more attribute lines
                }
                attributesEncountered = true;

                try
                {
                    var (nextStream, attr) = ParseSingleAttributeLine(currentStream); // currentStream includes leading spaces for the line
                    currentStream = nextStream;

                    string key = attr.Key;
                    string value = attr.Value;

                    if (key.Equals("VB_Name", StringComparison.OrdinalIgnoreCase))
                    {
                        attributes.Name = value;
                        nameFound = true;
                    }
                    else if (key.Equals("VB_GlobalNameSpace", StringComparison.OrdinalIgnoreCase))
                        attributes.GlobalNameSpace = (value.Equals("True", StringComparison.OrdinalIgnoreCase) || value == "-1" || value == "1") ? NameSpaceKind.Global : NameSpaceKind.Local;
                    else if (key.Equals("VB_Creatable", StringComparison.OrdinalIgnoreCase))
                        attributes.Creatable = (value.Equals("True", StringComparison.OrdinalIgnoreCase) || value == "-1" || value == "1") ? CreatableKind.True : CreatableKind.False;
                    else if (key.Equals("VB_PredeclaredId", StringComparison.OrdinalIgnoreCase))
                        attributes.PreDeclaredID = (value.Equals("True", StringComparison.OrdinalIgnoreCase) || value == "-1" || value == "1") ? PreDeclaredIDKind.True : PreDeclaredIDKind.False;
                    else if (key.Equals("VB_Exposed", StringComparison.OrdinalIgnoreCase))
                        attributes.Exposed = (value.Equals("True", StringComparison.OrdinalIgnoreCase) || value == "-1" || value == "1") ? ExposedKind.True : ExposedKind.False;
                    else if (key.Equals("VB_Description", StringComparison.OrdinalIgnoreCase))
                        attributes.Description = value;
                    else if (key.Equals("VB_Ext_KEY", StringComparison.OrdinalIgnoreCase))
                    {
                        // VB_Ext_KEY = "KeyString", "ValueString"
                        // The value from ParseSingleAttributeLine will be "\"KeyString\", \"ValueString\""
                        // We need to parse this composite value.
                        var parts = value.Split(new[] { "\", \"" }, StringSplitOptions.None);
                        if (parts.Length == 2)
                        {
                            string extKeyName = parts[0].StartsWith("\"") ? parts[0].Substring(1) : parts[0];
                            string extKeyValue = parts[1].EndsWith("\"") ? parts[1].Substring(0, parts[1].Length - 1) : parts[1];
                            attributes.ExtKey[extKeyName] = extKeyValue;
                        } else {
                             // Log warning or throw if format is unexpected
                             attributes.ExtKey["_MalformedExtKey_" + Guid.NewGuid().ToString("N").Substring(0,6)] = value;
                        }
                    }
                    else // Other potential attributes not specifically listed in VB6FileAttributes struct
                    {
                         attributes.ExtKey[key] = value; // Store in ExtKey as a fallback
                    }
                }
                catch (VB6ParseException ex)
                {
                    // If an attribute line starts but is malformed, it's an error.
                    throw currentStream.CreateExceptionFromCheckpoint(lineStartCheckpoint,VB6ErrorKind.AttributeParseError, innerException: ex);
                }
            }
            
            // Check if VB_Name was found if any attributes were processed
            if (attributesEncountered && !nameFound)
            {
                throw stream.CreateException(VB6ErrorKind.MissingNameAttribute, 
                    innerException: new Exception("An attribute block was parsed, but it did not contain the mandatory 'VB_Name' attribute."));
            }

            return (currentStream, attributes);
        }
        
        /// <summary>
        /// Parses a simple "Key = Value" line, where value can be quoted or unquoted.
        /// Consumes the line ending.
        /// </summary>
        public static (VB6Stream NewStream, string Key, string Value) ParseKeyValueLine(VB6Stream stream, char separator = '=')
        {
            var initialLineCheckpoint = stream.Checkpoint();
            VB6Stream s = SkipSpace0(stream);

            var (sAfterKey, keyStr) = TakeUntilAny(s, (byte)separator, (byte)' ', (byte)'\t');
            string key = keyStr.Trim();
            if (string.IsNullOrEmpty(key))
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.NameUnparsable, new Exception("Property key is missing or empty."));

            s = SkipSpace0(sAfterKey);
            if (s.IsEmpty || s.PeekByte() != (byte)separator)
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.NoKeyValueDividerFound, innerException: new Exception($"Missing separator '{separator}' after key '{key}'."));
            
            s = SkipSpace0(s.Advance(1)); // Consume separator and spaces

            string valueStr;
            if (!s.IsEmpty && s.PeekByte() == (byte)'"')
            {
                var (sAfterStringValue, val) = VB6.ParseVB6String(s); 
                valueStr = val;
                s = sAfterStringValue;
            }
            else
            {
                var (sAfterValue, val) = TakeUntilLineEndingOrComment(s);
                valueStr = val.Trim();
                s = sAfterValue;
            }
            
            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, key, valueStr);
        }
        
        /// <summary>
        /// Parses a property line linking to a resource: "Key = "ResourceFile.frx":OffsetHex"
        /// Consumes the line ending.
        /// </summary>
        public static (VB6Stream NewStream, string Key, string ResourceFile, uint Offset) ParseKeyResourceOffsetLine(VB6Stream stream)
        {
            var initialLineCheckpoint = stream.Checkpoint();
            VB6Stream s = SkipSpace0(stream);

            var (sAfterKey, keyStr) = TakeUntilAny(s, " \t="u8.ToArray());
            string key = keyStr.Trim();
            if (string.IsNullOrEmpty(key))
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.NameUnparsable, new Exception("Property key for resource link is missing."));
            
            s = SkipSpace0(sAfterKey);
            if (s.IsEmpty || s.PeekByte() != (byte)'=')
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.NoEqualSplit, innerException: new Exception($"Missing '=' after resource link key '{key}'."));
            
            s = SkipSpace0(s.Advance(1));

            if (s.IsEmpty || s.PeekByte() != (byte)'"')
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.ResourceFileNameUnparsable, new Exception("Resource file name not enclosed in quotes."));

            var (sAfterFileName, fileName) = VB6.ParseVB6String(s); // Handles internal quotes if any
            s = sAfterFileName;
            s = SkipSpace0(s); // NEW

            if (s.IsEmpty || s.PeekByte() != (byte)':')
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.NoColonForOffsetSplit, innerException: new Exception("Missing ':' separator for resource offset."));
            
            s = s.Advance(1); // Consume ':'
            s = SkipSpace0(s); // NEW

            var (sAfterOffset, offsetStrHex) = VB6.TakeWhile(s, b => 
                (b >= (byte)'0' && b <= (byte)'9') ||
                (b >= (byte)'A' && b <= (byte)'F') ||
                (b >= (byte)'a' && b <= (byte)'f')
            );
            s = sAfterOffset;
            
            if (string.IsNullOrEmpty(offsetStrHex) || !uint.TryParse(offsetStrHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint offset))
                throw s.CreateExceptionFromCheckpoint(initialLineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.OffsetUnparsable, new Exception($"Invalid or missing hex offset: '{offsetStrHex}'."));

            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, key, fileName, offset);
        }

        /// <summary>
        /// Parses "Object=" lines, which can refer to compiled components (GUID) or sub-projects (path).
        /// This is used in .VBP project files and sometimes in .FRM headers for ActiveX controls.
        /// </summary>
        public static (VB6Stream NewStream, VB6ObjectReference ParsedReference) ParseObjectLine(VB6Stream stream, string originalLineForErrorContext)
        {
            var initialLineStartCheckpoint = stream.Checkpoint();
            VB6Stream s = stream;
            var objRef = new VB6ObjectReference { OriginalValue = originalLineForErrorContext };
            var contentStartCheckpoint = s.Checkpoint();

            // 1. Check for Sub-Project Path (starts with '*')
            if (!s.IsEmpty && s.PeekByte() == (byte)'*')
            {
                // Check for quoted subproject: "*\A..." or "*\\A..." inside quotes
                if (s.Length > 1 && s.PeekByteAt(1) == (byte)'"' && s.Length > 5 &&
                    (s.PeekByteAt(2) == (byte)'*' && (s.PeekByteAt(3) == (byte)'\\' || s.PeekByteAt(3) == (byte)'/') && s.PeekByteAt(4) == (byte)'A'))
                { // This pattern seems unlikely: Object=*"*\A..." - more likely Object="*\A..."
                  // Let's assume the simpler Object="*\A..." or Object=*\A...
                    throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.UnrecognizedObjectFormat, innerException: new Exception("Unexpected quoted sub-project format."));
                }
                else if (!s.IsEmpty && s.PeekByte() == (byte)'"' && s.Length > 4 && s.PeekByteAt(1) == (byte)'*' &&
                         (s.PeekByteAt(2) == (byte)'\\' || s.PeekByteAt(2) == (byte)'/') && s.PeekByteAt(3) == (byte)'A')
                { // Format: Object="*\AProjectPath"
                    s = s.Advance(1); // Consume leading quote
                    var (sAfterPath, pathContent) = VB6.TakeUntil(s, (byte)'"');
                    objRef.Path = pathContent; // Path *is* the content inside quotes
                    s = sAfterPath;
                    if (s.IsEmpty || s.PeekByte() != (byte)'"')
                        throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.ExpectedClosingQuote, innerException: new Exception("Missing closing quote for quoted project object path."));
                    s = s.Advance(1); // Consume closing quote
                }
                else if (!s.IsEmpty && s.PeekByte() == (byte)'*' && s.Length > 2 &&
                         ((s.PeekByteAt(1) == (byte)'\\' || s.PeekByteAt(1) == (byte)'/') && s.PeekByteAt(2) == (byte)'A'))
                { // Format: Object=*\AProjectPath
                    var (sAfterPath, pathContent) = TakeUntilLineEndingOrComment(s);
                    objRef.Path = pathContent.Trim();
                    s = sAfterPath;
                }
                else
                {
                    throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.UnrecognizedObjectFormat, innerException: new Exception("Malformed path-based object reference."));
                }
            }
            // 2. GUID-based Compiled Object
            else
            {
                string guidBlockString;
                VB6Stream sAfterGuidBlockParsing;

                // Check if the GUID block itself is quoted: Object="{GUID_PART}"; "FILE.OCX"
                if (!s.IsEmpty && s.PeekByte() == (byte)'"')
                {
                    var (tempS, contentInsideQuotes) = VB6.ParseVB6String(s); // Consumes outer quotes
                    guidBlockString = contentInsideQuotes;
                    sAfterGuidBlockParsing = tempS;
                }
                else // GUID block is not quoted: Object={GUID_PART}; "FILE.OCX"
                {
                    var (tempS, content) = TakeUntilAny(s, (byte)';', (byte)'\r', (byte)'\n', (byte)'\'');
                    guidBlockString = content.Trim();
                    sAfterGuidBlockParsing = tempS;
                }

                var guidParts = guidBlockString.Split('#');
                if (guidParts.Length < 1)
                    throw s.CreateExceptionFromCheckpoint(initialLineStartCheckpoint, VB6ErrorKind.UnrecognizedObjectFormat, innerException: new Exception("Object reference GUID part is missing or malformed."));

                string guidStrToParse = guidParts[0].Trim();
                if (guidStrToParse.StartsWith("{") && guidStrToParse.EndsWith("}"))
                    guidStrToParse = guidStrToParse.Substring(1, guidStrToParse.Length - 2);

                if (!Guid.TryParse(guidStrToParse, out Guid parsedGuid))
                    throw s.CreateExceptionFromCheckpoint(initialLineStartCheckpoint, VB6ErrorKind.UnableToParseUuid, innerException: new Exception($"Invalid GUID string: {guidParts[0].Trim()}"));
                objRef.ObjectGuidString = parsedGuid.ToString("B").ToUpperInvariant();

                objRef.Version = guidParts.Length > 1 ? guidParts[1].Trim() : null;
                objRef.LocaleID = guidParts.Length > 2 ? guidParts[2].Trim() : null; // Renamed from LocaleID

                s = sAfterGuidBlockParsing;

                // Part 2: Optional Semicolon and FileName
                s = SkipSpace0(s);
                if (!s.IsEmpty && s.PeekByte() == (byte)';')
                {
                    s = SkipSpace0(s.Advance(1)); // Consume ';' and spaces

                    if (!s.IsEmpty && s.PeekByte() == (byte)'"') // Quoted Filename
                    {
                        var (sAfterFileName, fileNameContent) = VB6.ParseVB6String(s);
                        objRef.FileName = fileNameContent; // ParseVB6String returns unquoted content
                        s = sAfterFileName;
                    }
                    else if (!s.IsEmpty) // Unquoted Filename (till EOL/comment)
                    {
                        var (sAfterFileName, fileNameContent) = TakeUntilLineEndingOrComment(s);
                        objRef.FileName = fileNameContent.Trim();
                        s = sAfterFileName;
                    }
                }
                // If no semicolon, then no filename part is assumed for GUID-based objects.
            }

            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, objRef);
        }

        public static (VB6Stream NewStream, VB6ObjectReference ParsedReference) ParseObjectLine_OLD (VB6Stream stream, string originalLineForErrorContext)
        {
            var initialLineStartCheckpoint = stream.Checkpoint(); // Checkpoint for the entire "Object=..." line for context
            VB6Stream s = stream; // s will be advanced

            // s is expected to be positioned *after* "Object=" and any leading spaces after "=".
            // If "Object=" is still present, it means the caller didn't consume it.
            // For robustness, let's assume the caller has positioned `s` at the start of the value part.
            
            var contentStartCheckpoint = s.Checkpoint(); // Checkpoint at the actual start of the value
            var objRef = new VB6ObjectReference { OriginalValue = originalLineForErrorContext };

            // Determine if it's a quoted project path, unquoted project path, or GUID-based object.
            // 1. Quoted Project Path: Object="*\AProjectPath"
            if (!s.IsEmpty && s.PeekByte() == (byte)'"' && s.Length > 4 && 
                (s.PeekByteAt(1) == (byte)'*' && (s.PeekByteAt(2) == (byte)'\\' || s.PeekByteAt(2) == (byte)'/') && s.PeekByteAt(3) == (byte)'A'))
            {
                s = s.Advance(1); // Consume leading quote
                var (sAfterPath, pathContent) = VB6.TakeUntil(s, (byte)'"'); // Path until closing quote
                objRef.Path = pathContent; // The path here is the part *after* "*\A"
                s = sAfterPath;
                if (s.IsEmpty || s.PeekByte() != (byte)'"')
                    throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.ExpectedClosingQuote, innerException: new Exception("Missing closing quote for quoted project object path."));
                s = s.Advance(1); // Consume closing quote
            }
            // 2. Unquoted Project Path: Object=*\AProjectPath
            else if (!s.IsEmpty && s.PeekByte() == (byte)'*' && s.Length > 2 &&
                     ((s.PeekByteAt(1) == (byte)'\\' || s.PeekByteAt(1) == (byte)'/') && s.PeekByteAt(2) == (byte)'A'))
            {
                var (sAfterPath, pathContent) = TakeUntilLineEndingOrComment(s); // Path is rest of line
                objRef.Path = pathContent.Trim();
                s = sAfterPath;
            }
            // 3. GUID-based Compiled Object: Object={GUID}#Version#Unknown1;FileName or Object="{GUID}#Version#Unknown1;FileName"
            else
            {
                bool isGuidPartQuoted = !s.IsEmpty && s.PeekByte() == (byte)'"';
                VB6Stream contentToParse = s;
                if (isGuidPartQuoted)
                {
                    contentToParse = s.Advance(1); // Skip outer quote if present
                    if (contentToParse.IsEmpty || contentToParse.PeekByte() != (byte)'{') // Ensure it's actually a GUID after quote
                         throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.ExpectedOpeningBraceForGuid, innerException: new Exception("Expected '{' for GUID after opening quote in object reference."));
                }

                if (contentToParse.IsEmpty || contentToParse.PeekByte() != (byte)'{')
                    throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.ExpectedOpeningBraceForGuid, innerException: new Exception("Expected '{' for GUID in object reference."));
                
                contentToParse = VB6.ParseLiteral(contentToParse, (byte)'{');
                var (sAfterGuidStr, guidStr) = VB6.TakeUntil(contentToParse, (byte)'}');
                if (!Guid.TryParse(guidStr, out Guid parsedGuid))
                    throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.UnableToParseUuid, innerException: new Exception($"Invalid GUID string: {guidStr}"));
                objRef.ObjectGuidString = parsedGuid.ToString("B").ToUpper(); // Standard GUID format
                contentToParse = VB6.ParseLiteral(sAfterGuidStr, (byte)'}');

                // Version and LocaleID (LCID/Flags)
                contentToParse = VB6.ParseLiteral(contentToParse, (byte)'#');
                var (sAfterVer, verStr) = VB6.TakeUntil(contentToParse, (byte)'#');
                objRef.Version = verStr;
                contentToParse = VB6.ParseLiteral(sAfterVer, (byte)'#');
                
                var (sAfterUnk1, unk1Str) = VB6.TakeUntil(contentToParse, (byte)';');
                objRef.LocaleID = unk1Str;
                contentToParse = VB6.ParseLiteral(sAfterUnk1, (byte)';');
                
                contentToParse = SkipSpace0(contentToParse); // Skip space before FileName
                
                string fileNameStr;
                if (isGuidPartQuoted) // If the whole object line was quoted, filename is also inside these quotes
                {
                    var (sAfterFile, fileN) = VB6.TakeUntil(contentToParse, (byte)'"'); // Filename until overall closing quote
                    fileNameStr = fileN;
                    contentToParse = sAfterFile;
                    if(contentToParse.IsEmpty || contentToParse.PeekByte() != (byte)'"') 
                        throw s.CreateExceptionFromCheckpoint(contentStartCheckpoint, VB6ErrorKind.ExpectedClosingQuote, innerException: new Exception("Missing closing quote for quoted object reference."));
                    contentToParse = contentToParse.Advance(1); // Consume overall closing quote
                }
                else // Unquoted object line, filename is till EOL/comment
                {
                    var (sAfterFile, fileN) = TakeUntilLineEndingOrComment(contentToParse);
                    fileNameStr = fileN.Trim();
                    contentToParse = sAfterFile;
                }
                objRef.FileName = fileNameStr;
                s = contentToParse; // Update main stream 's'
            }

            s = ParseOptionalTrailingCommentAndLineEnding(s);
            return (s, objRef);
        }
    }
}