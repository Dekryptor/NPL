using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using VBSix.Errors;
using VBSix.Language;
using VBSix.Language.Controls;
using VBSix.Utilities;

namespace VBSix.Parsers
{
    /// <summary>
    /// Delegate for resolving resource files (.frx).
    /// </summary>
    /// <param name="frxFilePath">The full path to the .frx file.</param>
    /// <param name="offset">The offset of the resource data within the .frx file.</param>
    /// <returns>A byte array containing the resource data, or null if not found or an error occurs.</returns>
    public delegate byte[]? ResourceFileResolver(string frxFilePath, int offset);

    /// <summary>
    /// Represents a property value within a VB6 .frm file.
    /// It can be a simple text string or a reference to binary data in an .frx file.
    /// </summary>
    public class PropertyValue
    {
        private readonly object _value; 

        public bool IsResource { get; }

        public PropertyValue(string textValue)
        {
            _value = textValue ?? string.Empty; 
            IsResource = false;
        }

        public PropertyValue(byte[] resourceValue)
        {
            _value = (byte[])(resourceValue ?? throw new ArgumentNullException(nameof(resourceValue))).Clone();
            IsResource = true;
        }

        public string? AsString() => IsResource ? null : (string)_value;
        public byte[]? AsResource() => IsResource ? (byte[])((byte[])_value).Clone() : null;
        public override string ToString() => IsResource ? $"[Binary Resource: {((byte[])_value).Length} bytes]" : (string)_value;
    }

    public class VB6FullyQualifiedName
    {
        public string Namespace { get; set; } = string.Empty;
        public string Kind { get; set; } = string.Empty; 
        public string Name { get; set; } = string.Empty;  

        public override string ToString() => $"{Namespace}.{Kind} {Name}";
    }

    public class VB6PropertyGroup
    {
        public string Name { get; set; } = string.Empty;
        public Guid? Guid { get; set; }
        public Dictionary<string, PropertyValue> Properties { get; set; } = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase);
        public List<VB6PropertyGroup> NestedGroups { get; set; } = [];
        public override string ToString() => $"PropertyGroup: {Name}" + (Guid.HasValue ? $" {{{Guid.Value}}}" : "");
    }
    
    public class ControlPropertiesDictionary 
    {
        public Dictionary<string, PropertyValue> Values { get; set; } = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase);
    }

    public partial class VB6FormFile
    {
        public VB6FileFormatVersion? FormatVersion { get; set; }
        public List<VB6ObjectReference> ObjectReferences { get; set; } = [];
        public VB6Control? Form { get; set; } 
        public VB6FileAttributes Attributes { get; set; } = new VB6FileAttributes();
        public List<VB6Token> Tokens { get; set; } = [];

        private static readonly Regex FrxLinkRegex = new Regex(@"^""?([^""\:]+\.(?:frx|ctx|pag|dsr|dsx))""?\s*:\s*([0-9A-Fa-f]+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        
        private static readonly Regex PropertyGroupGuidRegex = new Regex(@"\s*\{\s*([0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12})\s*\}$", RegexOptions.Compiled | RegexOptions.RightToLeft);


        public static VB6FormFile Parse(string formPath, byte[] inputBytes) => ParseWithResolver(formPath, inputBytes, DefaultResourceFileResolver);

        public static VB6FormFile ParseWithResolver(string formPath, byte[] inputBytes, ResourceFileResolver resourceResolver)
        {
            var formFile = new VB6FormFile();
            var stream = new VB6Stream(formPath, inputBytes);
            var initialStreamForError = stream.Checkpoint();

            try
            {
                (VB6Stream sAfterVersion, VB6FileFormatVersion version, HeaderKind _) = HeaderParser.ParseVersionLine(stream, HeaderKind.Form);
                formFile.FormatVersion = version;
                stream = sAfterVersion;

                (stream, formFile.ObjectReferences) = ParseFormObjects(stream);
                stream = SkipBlankLinesAndComments(stream);

                VB6FullyQualifiedName mainFqn = ParseBeginFormLine(ref stream);

                (VB6Stream streamAfterControlBlock, VB6Control mainControl) = ParseControlBlock(stream, mainFqn, resourceResolver, 0, formPath);
                formFile.Form = mainControl;
                stream = streamAfterControlBlock;

                stream = SkipBlankLinesAndComments(stream);

                (VB6Stream streamAfterAttrs, VB6FileAttributes attributes) = HeaderParser.ParseAttributes(stream);
                formFile.Attributes = attributes;
                stream = streamAfterAttrs;

                formFile.Tokens = VB6.Vb6Parse(stream);

                return formFile;
            }
            catch (VB6ParseException) { throw; }
            catch (Exception ex)
            {
                throw stream.CreateExceptionFromCheckpoint(initialStreamForError, VB6ErrorKind.Unparseable,
                    innerException: new Exception($"Unexpected error parsing form file '{formPath}'. Details: {ex.Message}", ex));
            }
        }
        
        public static VB6Stream SkipBlankLinesAndComments(VB6Stream stream)
        {
            VB6Stream current = stream;
            while (!current.IsEmpty)
            {
                VB6Stream temp = HeaderParser.SkipSpace0(current);
                if (temp.IsEmpty) break;

                byte? peek = temp.PeekByte();
                if (peek == (byte)'\'')
                {
                    (VB6Stream s, string _) = HeaderParser.TakeUntilLineEnding(temp);
                    current = HeaderParser.ParseLineEnding(s);
                }
                else if (peek == (byte)'\r' || peek == (byte)'\n')
                {
                    current = HeaderParser.ParseLineEnding(temp);
                }
                else
                {
                    return current;
                }
            }
            return current;
        }

        public static (VB6Stream NewStream, List<VB6ObjectReference> References) ParseFormObjects(VB6Stream stream)
        {
            var references = new List<VB6ObjectReference>();
            VB6Stream currentStream = stream;

            while (!currentStream.IsEmpty)
            {
                var lineStartCheckpoint = currentStream.Checkpoint();
                VB6Stream streamAtLineStart = currentStream;
                currentStream = HeaderParser.SkipSpace0(currentStream);

                if (currentStream.IsEmpty || !currentStream.CompareSliceCaseless("Object="))
                {
                    return (streamAtLineStart, references);
                }

                (VB6Stream _, string originalLineContent) = HeaderParser.TakeUntilLineEnding(streamAtLineStart);

                try
                {
                    (VB6Stream sAfterObjectEquals, string _) = VB6.TakeSpecificString(currentStream, "Object=", caseSensitive: true);
                    VB6Stream valuePartStream = HeaderParser.SkipSpace0(sAfterObjectEquals);

                    (VB6Stream nextStreamAfterObject, VB6ObjectReference parsedRef) = HeaderParser.ParseObjectLine(valuePartStream, originalLineContent.Trim());
                    references.Add(parsedRef);
                    currentStream = nextStreamAfterObject;
                }
                catch (VB6ParseException ex)
                {
                    throw stream.CreateExceptionFromCheckpoint(lineStartCheckpoint, ex.Kind, ex.PropertyErrorDetail, new Exception($"Error parsing Object line: '{originalLineContent.Trim()}'", ex));
                }
                catch (Exception ex)
                {
                    throw stream.CreateExceptionFromCheckpoint(lineStartCheckpoint, VB6ErrorKind.UnrecognizedObjectFormat, innerException: new Exception($"Generic error parsing Object line: '{originalLineContent.Trim()}'", ex));
                }
            }
            return (currentStream, references);
        }
        
        private static VB6FullyQualifiedName ParseBeginFormLine(ref VB6Stream stream)
        {
            var initialCheckpoint = stream.Checkpoint();
            VB6Stream s = HeaderParser.SkipSpace0(stream);

            if (!s.CompareSliceCaseless("BEGIN "))
                throw s.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.NoBegin, innerException: new Exception("Expected 'BEGIN ' keyword for main form/control block."));

            (VB6Stream sAfterBeginKeyword, string _) = VB6.TakeSpecificString(s, "BEGIN", caseSensitive: true);
            s = HeaderParser.SkipSpace1(sAfterBeginKeyword);

            var fqnResultTuple = ParseFullyQualifiedNameInternal(s);
            if (fqnResultTuple == null)
                throw s.CreateExceptionFromCheckpoint(initialCheckpoint, VB6ErrorKind.NoForm, innerException: new Exception("Could not parse form's fully qualified name after BEGIN."));

            stream = fqnResultTuple.Value.stream;
            return fqnResultTuple.Value.fqn;
        }

        public static (VB6Stream stream, VB6FullyQualifiedName fqn)? ParseFullyQualifiedNameInternal(VB6Stream stream)
        {
            var initialFqnCheckpoint = stream.Checkpoint();
            VB6Stream s = stream;

            (VB6Stream sAfterNamespace, string namespaceStr) = HeaderParser.TakeUntil(s, (byte)'.');
            string ns = namespaceStr.Trim();
            if (string.IsNullOrEmpty(ns))
                throw s.CreateExceptionFromCheckpoint(initialFqnCheckpoint, VB6ErrorKind.NoNamespaceAfterBegin);

            if (sAfterNamespace.IsEmpty || sAfterNamespace.PeekByte() != (byte)'.')
                throw s.CreateExceptionFromCheckpoint(initialFqnCheckpoint, VB6ErrorKind.NoDotAfterNamespace, innerException: new Exception($"Expected '.' after namespace '{ns}'."));
            s = sAfterNamespace.Advance(1);

            (VB6Stream sAfterKind, string kindStr) = HeaderParser.TakeUntilAny(s, " \t"u8.ToArray());
            string kind = kindStr.Trim();
            if (string.IsNullOrEmpty(kind))
                throw s.CreateExceptionFromCheckpoint(initialFqnCheckpoint, VB6ErrorKind.NoUserControlNameAfterDot, innerException: new Exception("Control kind is missing."));

            s = HeaderParser.SkipSpace1(sAfterKind);

            (VB6Stream sAfterName, string nameStr) = HeaderParser.TakeUntilLineEndingOrComment(s);
            string ctrlName = nameStr.Trim();
            if (string.IsNullOrEmpty(ctrlName))
                throw s.CreateExceptionFromCheckpoint(initialFqnCheckpoint, VB6ErrorKind.NoControlNameAfterControlKind);

            s = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(sAfterName);

            var fqn = new VB6FullyQualifiedName { Namespace = ns, Kind = kind, Name = ctrlName };
            return (s, fqn);
        }
        
        private static bool TryParseFrxLink(string value, out string? frxFileName, out int frxOffset)
        {
            frxFileName = null;
            frxOffset = 0;
            Match match = FrxLinkRegex.Match(value);
            if (match.Success)
            {
                frxFileName = match.Groups[1].Value;
                // Group 2 is the hex offset string
                if (int.TryParse(match.Groups[2].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out frxOffset))
                {
                    return true;
                }
                else
                {
                    // Invalid hex offset string
                    // Console.Error.WriteLine($"Warning: Invalid FRX offset format '{match.Groups[2].Value}' in '{value}'");
                    return false; 
                }
            }
            return false;
        }

        /// <summary>
        /// Default resource resolver implementation.
        /// (Detailed implementation from prompt goes here)
        /// </summary>
        public static byte[]? DefaultResourceFileResolver(string frxFilePath, int offset)
        {
            byte[] buffer;
            try
            {
                if (!File.Exists(frxFilePath)) return null;
                buffer = File.ReadAllBytes(frxFilePath);
            }
            catch (IOException) { return null; }

            if (offset < 0 || offset >= buffer.Length) return null;

            byte[]? ReadSlice(int start, int length)
            {
                if (start < 0 || length < 0 || start + length > buffer.Length) return null;
                byte[] slice = new byte[length];
                Array.Copy(buffer, start, slice, 0, length);
                return slice;
            }

            if (offset + 8 <= buffer.Length)
            {
                byte[]? sig = ReadSlice(offset + 4, 4);
                if (sig != null && sig.SequenceEqual(new byte[] { (byte)'l', (byte)'t', 0, 0 }))
                {
                    if (offset + 12 > buffer.Length) return null;
                    byte[]? size1Bytes = ReadSlice(offset, 4);
                    byte[]? size2Bytes = ReadSlice(offset + 8, 4);
                    if (size1Bytes == null || size2Bytes == null) return null;

                    uint size1 = BitConverter.ToUInt32(size1Bytes, 0);
                    uint size2 = BitConverter.ToUInt32(size2Bytes, 0);

                    if (size1 == 8 && size2 == 0) return [];
                    if (size2 != size1 - 8) return null;

                    return ReadSlice(offset + 12, (int)size2);
                }
            }

            if (buffer[offset] == 0xFF)
            {
                if (offset + 3 > buffer.Length) return null;
                byte[]? lenBytes = ReadSlice(offset + 1, 2);
                if (lenBytes == null) return null;
                int recordLen = BitConverter.ToUInt16(lenBytes, 0);

                if (offset + 3 + recordLen > buffer.Length && recordLen > 0)
                {
                    recordLen--;
                }
                return ReadSlice(offset + 3, recordLen);
            }

            if (offset + 4 <= buffer.Length)
            {
                byte[]? listSig = ReadSlice(offset + 2, 2);
                bool isList = listSig != null &&
                              (listSig.SequenceEqual(new byte[] { 0x03, 0x00 }) ||
                               listSig.SequenceEqual(new byte[] { 0x07, 0x00 }));
                if (isList)
                {
                    if (offset + 2 > buffer.Length) return null;
                    byte[]? countBytes = ReadSlice(offset, 2);
                    if (countBytes == null) return null;
                    ushort itemCount = BitConverter.ToUInt16(countBytes, 0);

                    int currentParseOffset = offset + 4;
                    for (int i = 0; i < itemCount; i++)
                    {
                        if (currentParseOffset + 2 > buffer.Length) return null;
                        byte[]? itemLenBytes = ReadSlice(currentParseOffset, 2);
                        if (itemLenBytes == null) return null;
                        ushort itemLen = BitConverter.ToUInt16(itemLenBytes, 0);
                        currentParseOffset += 2 + itemLen;
                        if (currentParseOffset > buffer.Length) return null;
                    }
                    return ReadSlice(offset, currentParseOffset - offset);
                }
            }

            if (buffer.Length >= 12 && offset + 4 <= buffer.Length)
            {
                bool hasNull = false;
                for (int k = 0; k < 4; ++k) if (buffer[offset + k] == 0) { hasNull = true; break; }
                if (hasNull)
                {
                    byte[]? sizeBytes = ReadSlice(offset, 4);
                    if (sizeBytes == null) return null;
                    int recordLen = (int)BitConverter.ToUInt32(sizeBytes, 0);
                    return ReadSlice(offset + 4, recordLen);
                }
            }

            {
                int recordLen = buffer[offset];
                if (offset + 1 + recordLen > buffer.Length && recordLen > 0)
                {
                    recordLen--;
                }
                if (recordLen < 0) return null;
                return ReadSlice(offset + 1, recordLen);
            }
        }

        public static List<string> ListResolver(byte[] listResourcePayload)
        {
            var listItems = new List<string>();
            if (listResourcePayload == null || listResourcePayload.Length < 2) return listItems;
            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            try
            {
                using var ms = new MemoryStream(listResourcePayload);
                using var br = new BinaryReader(ms, registeredEncoding);
                if (ms.Length < 2) return listItems;
                ushort itemCount = br.ReadUInt16();
                if (ms.Length < 4) return listItems;
                ms.Seek(2, SeekOrigin.Current);

                for (int i = 0; i < itemCount; i++)
                {
                    if (ms.Position + 2 > ms.Length) break;
                    ushort itemLen = br.ReadUInt16();
                    if (ms.Position + itemLen > ms.Length) break;
                    byte[] itemBytes = br.ReadBytes(itemLen);
                    listItems.Add(registeredEncoding.GetString(itemBytes));
                }
            }
            catch { /* Optional: Log error */ }
            return listItems;
        }

        public static List<int> ItemDataResolver(byte[] itemDataResourcePayload)
        {
            var itemDataList = new List<int>();
            if (itemDataResourcePayload == null || itemDataResourcePayload.Length == 0) return itemDataList;
            try
            {
                using var ms = new MemoryStream(itemDataResourcePayload);
                using var br = new BinaryReader(ms);
                while (ms.Position + 4 <= ms.Length)
                {
                    itemDataList.Add(br.ReadInt32());
                }
            }
            catch { /* Optional: Log error */ }
            return itemDataList;
        }

        public static List<OleVerbInfo> OleVerbResolver(byte[] verbsPayload)
        {
            var verbs = new List<OleVerbInfo>();
            if (verbsPayload == null || verbsPayload.Length < 2) return verbs;
            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            try
            {
                using var ms = new MemoryStream(verbsPayload);
                using var br = new BinaryReader(ms, registeredEncoding);
                if (ms.Length < 2) return verbs;
                ushort verbCount = br.ReadUInt16();
                for (int i = 0; i < verbCount; i++)
                {
                    if (ms.Position + 4 + 4 + 2 > ms.Length) break;
                    int verbId = br.ReadInt32();
                    int verbFlags = br.ReadInt32();
                    ushort nameLength = br.ReadUInt16();
                    if (ms.Position + nameLength > ms.Length) break;
                    byte[] nameBytes = br.ReadBytes(nameLength);
                    string verbName = registeredEncoding.GetString(nameBytes);
                    verbs.Add(new OleVerbInfo { Id = verbId, Name = verbName, Flags = verbFlags });
                }
            }
            catch { /* Log error */ }
            return verbs;
        }

        public static (VB6Stream stream, VB6Control control) ParseControlBlock(
            VB6Stream stream,
            VB6FullyQualifiedName fqn,
            ResourceFileResolver resourceResolver,
            int depth,
            string formFilePath)
        {
            var rawProperties = new ControlPropertiesDictionary();
            var propertyGroups = new List<VB6PropertyGroup>();
            var subControls = new List<VB6Control>();
            VB6Stream currentStream = stream;
            string parentDirectory = Path.GetDirectoryName(formFilePath) ?? string.Empty;

            while (!currentStream.IsEmpty)
            {
                var lineStartCheckpoint = currentStream.Checkpoint();
                VB6Stream streamForThisLine = currentStream;
                currentStream = HeaderParser.SkipSpace0(currentStream);

                if (currentStream.IsEmpty)
                    throw stream.CreateExceptionFromCheckpoint(lineStartCheckpoint, VB6ErrorKind.NoEndKeyword, innerException: new Exception($"Unexpected end of stream while parsing control block for '{fqn.Name}'. Expected 'END'."));

                if (currentStream.CompareSliceCaseless("END"))
                {
                    (VB6Stream s_end, string _) = VB6.TakeSpecificString(currentStream, "END", caseSensitive: true);
                    s_end = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(s_end);
                    currentStream = s_end;
                    VB6Control builtControl = BuildControl(fqn, rawProperties, propertyGroups, subControls, formFilePath, resourceResolver);
                    return (currentStream, builtControl);
                }
                if (currentStream.CompareSliceCaseless("BEGIN "))
                {
                    VB6Stream s_begin = currentStream;
                    VB6FullyQualifiedName nestedFqn = ParseBeginFormLine(ref s_begin); // s_begin is updated by ref

                    (VB6Stream streamAfterNestedBlock, VB6Control nestedControl) = ParseControlBlock(s_begin, nestedFqn, resourceResolver, depth + 1, formFilePath);
                    subControls.Add(nestedControl);
                    currentStream = streamAfterNestedBlock;
                    continue;
                }
                if (currentStream.CompareSliceCaseless("BeginProperty "))
                {
                    (VB6Stream streamAfterPropertyGroup, VB6PropertyGroup propertyGroup) = ParsePropertyGroup(currentStream, resourceResolver, depth + 1, formFilePath);
                    propertyGroups.Add(propertyGroup);
                    currentStream = streamAfterPropertyGroup;
                    continue;
                }

                try
                {
                    (VB6Stream sAfterPropParse, string key, PropertyValue propValObject) = ParsePropertyKeyValue(streamForThisLine, parentDirectory, resourceResolver);
                    rawProperties.Values[key] = propValObject;
                    currentStream = sAfterPropParse;
                }
                catch (VB6ParseException) { throw; }
                catch (Exception ex)
                {
                    (VB6Stream _, string originalLineForError) = HeaderParser.TakeUntilLineEnding(streamForThisLine);
                    throw streamForThisLine.CreateExceptionFromCheckpoint(lineStartCheckpoint, VB6ErrorKind.KeyValueParseError,
                        innerException: new Exception($"Error parsing property line '{originalLineForError.Trim()}' for control '{fqn.Name}'. Details: {ex.Message}", ex));
                }
            }
            throw stream.CreateExceptionFromCheckpoint(stream.Checkpoint(), VB6ErrorKind.NoEndKeyword, innerException: new Exception($"Missing 'END' for control block '{fqn.Name}'."));
        }

        /// <summary>
        /// Tries to parse a string value as a link to an FRX resource.
        /// Handles standard "FileName.frx":OffsetHex and RichTextBox-specific $"FileName.frx":OffsetHex formats.
        /// </summary>
        /// <param name="value">The string value to parse.</param>
        /// <param name="frxFileName">Outputs the name of the .frx file if parsing is successful.</param>
        /// <param name="frxOffset">Outputs the hexadecimal offset within the .frx file if parsing is successful.</param>
        /// <returns>True if the value is a valid FRX link and parsed successfully; otherwise, false.</returns>
        internal static bool TryParseFrxLinkInternal(string value, out string? frxFileName, out int frxOffset)
        {
            frxFileName = null;
            frxOffset = 0;
            if (string.IsNullOrWhiteSpace(value)) return false;

            // Handle the optional '$' prefix commonly seen with RichTextBox TextRTF properties.
            // Example: $"Bar.frx":44A8D  or  "Bar.frx":0000
            string effectiveValue = value;
            if (value.Length > 2 && value.StartsWith("$\"") && value.EndsWith("\""))
            {
                effectiveValue = value[1..]; // Remove leading '$', keep surrounding quotes for regex
            }

            // Regex updated to be more specific about extensions and capture groups.
            // It expects:
            // 1. Optional quote.
            // 2. Filename (group "filename") ending with .frx, .ctx, .pag, .dsr, or .dsx.
            // 3. Optional quote.
            // 4. Colon separator, possibly surrounded by spaces.
            // 5. Hexadecimal offset (group "offset").
            const string frxPattern = @"^""?                                     # Optional opening quote for filename
                                        (?<filename>[^""\:]+\.(?:frx|ctx|pag|dsr|dsx)) # Filename (non-colon chars) + specific extensions
                                        ""?\s*:\s*                                   # Optional closing quote, colon, optional spaces
                                        (?<offset>[0-9A-Fa-f]+)                      # Hexadecimal offset
                                        $";

            Match match = Regex.Match(effectiveValue, frxPattern, RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.ExplicitCapture);

            if (match.Success)
            {
                frxFileName = match.Groups["filename"].Value;
                string offsetHexStr = match.Groups["offset"].Value;

                if (int.TryParse(offsetHexStr, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out frxOffset))
                {
                    return true;
                }
                else
                {
                    // This indicates a malformed hex offset string, though the regex should ensure it's hex digits.
                    // Could log a warning if strict error reporting on malformed (but matched regex) hex is needed.
                    // Console.Error.WriteLine($"Warning: Invalid hex offset '{offsetHexStr}' in FRX link: '{value}'");
                    frxFileName = null; // Ensure out params are reset on failure
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// Parses a single "Key = Value" property line from a control's definition block.
        /// The value can be a simple string, a number, or a link to an FRX resource.
        /// This method consumes the entire line including its line ending.
        /// </summary>
        private static (VB6Stream stream, string key, PropertyValue value) ParsePropertyKeyValue(
            VB6Stream stream, string parentFormDirectory, ResourceFileResolver resourceResolver)
        {
            VB6Stream currentStream = stream;
            var initialPropLineCheckpoint = currentStream.Checkpoint(); // For error reporting context

            currentStream = HeaderParser.SkipSpace0(currentStream);
            if (currentStream.IsEmpty)
            {
                // This case should ideally be caught by the caller (ParseControlBlock) checking for "END" or other keywords.
                // If reached, it implies an unexpected empty line where a property or keyword was expected.
                throw currentStream.CreateExceptionFromCheckpoint(initialPropLineCheckpoint, VB6ErrorKind.KeyValueParseError,
                    innerException: new Exception("Property line is empty or consists only of whitespace after skipping leading spaces."));
            }

            // 1. Parse Property Key
            // Key ends at the first space, tab, or '='.
            (VB6Stream sAfterKeyCandidate, string keyStr) = HeaderParser.TakeUntilAny(currentStream, "= \t"u8.ToArray());
            string key = keyStr.Trim();
            if (string.IsNullOrEmpty(key))
            {
                throw currentStream.CreateExceptionFromCheckpoint(initialPropLineCheckpoint, VB6ErrorKind.NameUnparsable, PropertyErrorType.NameUnparsable,
                    new Exception("Property name is missing or empty."));
            }
            currentStream = sAfterKeyCandidate;

            // 2. Parse "=" Separator
            currentStream = HeaderParser.SkipSpace0(currentStream); // Skip spaces before '='
            if (currentStream.IsEmpty || currentStream.PeekByte() != (byte)'=')
            {
                throw currentStream.CreateExceptionFromCheckpoint(initialPropLineCheckpoint, VB6ErrorKind.NoEqualSplit,
                    innerException: new Exception($"Missing '=' separator after property name '{key}'."));
            }
            currentStream = HeaderParser.SkipSpace0(currentStream.Advance(1)); // Consume '=' and skip spaces after it

            // 3. Parse Property Value
            PropertyValue propValue;
            VB6Stream streamAfterValueParsing;
            string rawValueStringForFrxCheck; // This will hold the string as it appears, e.g., "foo.frx:0000" or $"foo.frx:0000" or "TextValue"

            if (!currentStream.IsEmpty && currentStream.PeekByte() == (byte)'"') // Value is a quoted string
            {
                _ = currentStream.Checkpoint();
                // Vb6.ParseVB6String consumes the outer quotes and returns the inner content, handling "" escapes.
                (VB6Stream sAfterStringContent, string stringContent) = VB6.ParseVB6String(currentStream);
                streamAfterValueParsing = sAfterStringContent;
                // For FRX check, reconstruct the original quoted string as TryParseFrxLinkInternal expects it with quotes
                _ = "\"" + stringContent.Replace("\"", "\"\"") + "\""; // Re-escape internal quotes for accurate representation of original if needed
                                                                       // More simply, just take the original slice for FRX check before ParseVB6String consumes it.
                                                                       // Let's re-evaluate how rawValueStringForFrxCheck is obtained for quoted strings.
                                                                       // The goal is to get the part *exactly* as it was in the .frm before Vb6.ParseVB6String.

                // To get the original quoted string for FRX check:
                // We need the original slice from the opening quote to the closing quote.
                // Vb6.ParseVB6String already advanced past it. We need to capture it before or reconstruct.
                // A simpler way: Peek the original string if it starts with a quote, before detailed parsing.
                _ = currentStream.Index;
                _ = streamAfterValueParsing.Index; // This is *after* the closing quote
                // The original raw slice for the quoted string value is _fullStreamData from valueStartIndex to valueEndIndex
                // However, Vb6Stream doesn't expose _fullStreamData directly for slicing easily.
                // A pragmatic way: if it's a quoted string, the `rawValueStringForFrxCheck` should be the part
                // that `TryParseFrxLinkInternal` expects.

                // Let's assume `rawValueStringForFrxCheck` should be the string *as if it were unquoted by VB6 first for FRX check*
                // For "File.frx":Offset, it's `File.frx:Offset`
                // For $"File.frx":Offset, it's `$File.frx:Offset` (after outer quotes removed)
                // The `TryParseFrxLinkInternal` now handles the `$` and outer quotes.
                // So, `valContent` (the unquoted string from `ParseVB6String`) is NOT what TryParseFrxLinkInternal expects.
                // We need the string that was *between* the quotes if it was like TextRTF = $"foo.frx":123
                // The `valContent` would be `$foo.frx:123`. This is correct for `TryParseFrxLinkInternal`.
                rawValueStringForFrxCheck = stringContent; // Use the unquoted content for TryParseFrxLinkInternal.

                if (TryParseFrxLinkInternal(rawValueStringForFrxCheck, out string? frxFileName, out int frxOffset) && frxFileName != null)
                {
                    string frxFullPath = Path.Combine(parentFormDirectory, frxFileName);
                    byte[]? resourceData = resourceResolver(frxFullPath, frxOffset);
                    propValue = resourceData != null ? new PropertyValue(resourceData) : new PropertyValue(stringContent); // Fallback to unquoted string
                }
                else
                {
                    propValue = new PropertyValue(stringContent); // Store unquoted string content
                }
            }
            else // Value is unquoted (number, boolean literal, or unquoted identifier/string)
            {
                (VB6Stream sAfterUnquotedValue, string unquotedValueRaw) = HeaderParser.TakeUntilLineEndingOrComment(currentStream);
                streamAfterValueParsing = sAfterUnquotedValue;
                rawValueStringForFrxCheck = unquotedValueRaw.Trim(); // Use trimmed for FRX check

                if (TryParseFrxLinkInternal(rawValueStringForFrxCheck, out string? frxFileName, out int frxOffset) && frxFileName != null)
                {
                    string frxFullPath = Path.Combine(parentFormDirectory, frxFileName);
                    byte[]? resourceData = resourceResolver(frxFullPath, frxOffset);
                    propValue = resourceData != null ? new PropertyValue(resourceData) : new PropertyValue(rawValueStringForFrxCheck);
                }
                else
                {
                    propValue = new PropertyValue(rawValueStringForFrxCheck);
                }
            }

            // 4. Consume optional trailing comment and line ending
            VB6Stream finalStream = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(streamAfterValueParsing);
            return (finalStream, key, propValue);
        }

        private static (VB6Stream stream, VB6PropertyGroup propertyGroup) ParsePropertyGroup(
            VB6Stream stream, ResourceFileResolver resourceResolver, int depth, string formFilePath)
        {
            var initialGroupCheckpoint = stream.Checkpoint();
            VB6Stream currentStream = HeaderParser.SkipSpace0(stream);
            var propertyGroup = new VB6PropertyGroup();

            if (!currentStream.CompareSliceCaseless("BeginProperty "))
                throw currentStream.CreateExceptionFromCheckpoint(initialGroupCheckpoint, VB6ErrorKind.KeywordNotFound, innerException: new Exception("Expected 'BeginProperty '."));

            (VB6Stream sAfterKeyword, string _) = VB6.TakeSpecificString(currentStream, "BeginProperty", caseSensitive: true);
            VB6Stream sNameLine = HeaderParser.SkipSpace1(sAfterKeyword);

            (VB6Stream sAfterNameLineContent, string nameLineContent) = HeaderParser.TakeUntilLineEndingOrComment(sNameLine);
            string fullLineTrimmed = nameLineContent.Trim();
            currentStream = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(sAfterNameLineContent);

            Match guidMatch = PropertyGroupGuidRegex.Match(fullLineTrimmed);
            if (guidMatch.Success && guidMatch.Groups[1].Success)
            {
                propertyGroup.Guid = Guid.Parse(guidMatch.Groups[1].Value);
                propertyGroup.Name = fullLineTrimmed[..guidMatch.Index].Trim();
            }
            else
            {
                propertyGroup.Name = fullLineTrimmed;
            }

            if (string.IsNullOrEmpty(propertyGroup.Name))
                throw stream.CreateExceptionFromCheckpoint(initialGroupCheckpoint, VB6ErrorKind.NoPropertyName);

            string parentDirectory = Path.GetDirectoryName(formFilePath) ?? string.Empty;

            while (!currentStream.IsEmpty)
            {
                var lineCheckpoint = currentStream.Checkpoint();
                VB6Stream tempStream = HeaderParser.SkipSpace0(currentStream);

                if (tempStream.CompareSliceCaseless("EndProperty"))
                {
                    (VB6Stream sEnd, string _) = VB6.TakeSpecificString(tempStream, "EndProperty", caseSensitive: true);
                    currentStream = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(sEnd);
                    return (currentStream, propertyGroup);
                }
                if (tempStream.CompareSliceCaseless("BeginProperty "))
                {
                    var (sAfterNested, nestedGroup) = ParsePropertyGroup(tempStream, resourceResolver, depth + 1, formFilePath);
                    propertyGroup.NestedGroups.Add(nestedGroup);
                    currentStream = sAfterNested;
                    continue;
                }

                (VB6Stream sAfterProp, string key, PropertyValue value) = ParsePropertyKeyValue(currentStream, parentDirectory, resourceResolver);
                propertyGroup.Properties[key] = value;
                currentStream = sAfterProp;
            }
            throw stream.CreateExceptionFromCheckpoint(initialGroupCheckpoint, VB6ErrorKind.NoEndProperty, innerException: new Exception($"Missing EndProperty for group '{propertyGroup.Name}'."));
        }

        /// <summary>
        /// Constructs a specific VB6Control object (Form, TextBox, CommandButton, etc.)
        /// from its parsed fully qualified name, raw properties, property groups, and sub-controls.
        /// This method determines the control's 'Kind' and populates its typed properties.
        /// </summary>
        private static VB6Control BuildControl(
            VB6FullyQualifiedName fqn,
            ControlPropertiesDictionary rawPropertiesContainer, // Contains rawProps.Values
            List<VB6PropertyGroup> propertyGroupsFromParser,
            List<VB6Control> subControlsFromParser,
            string formFilePath, // Full path to the .frm/.ctl file for resolving relative FRX paths
            ResourceFileResolver resourceResolver // Delegate to resolve .frx data
            )
        {
            var control = new VB6Control { Name = fqn.Name };
            var rawProps = rawPropertiesContainer.Values; // Get the IDictionary<string, PropertyValue>

            // Populate common control properties
            control.Tag = rawProps.GetString("Tag", string.Empty);
            control.Index = rawProps.GetInt32("Index", 0); // Default to 0 if not specified

            // Determine the specific control kind and instantiate it
            if (fqn.Namespace.Equals("VB", StringComparison.OrdinalIgnoreCase))
            {
                switch (fqn.Kind)
                {
                    case "Form":
                        control.Kind = new VB6FormKind(rawProps, propertyGroupsFromParser, [.. subControlsFromParser.Where(sc => sc.Kind is not VB6MenuKind)], [.. subControlsFromParser.Where(sc => sc.Kind is VB6MenuKind).Select(ConvertVB6ControlToMenuControl)]);
                        break;
                    case "MDIForm":
                        control.Kind = new VB6MDIFormKind(rawProps, propertyGroupsFromParser, [.. subControlsFromParser.Where(sc => sc.Kind is not VB6MenuKind)], [.. subControlsFromParser.Where(sc => sc.Kind is VB6MenuKind).Select(ConvertVB6ControlToMenuControl)]);
                        break;
                    case "Menu":
                        control.Kind = new VB6MenuKind(rawProps, propertyGroupsFromParser, [.. subControlsFromParser.Select(ConvertVB6ControlToMenuControl)]);
                        break;
                    case "Frame":
                        control.Kind = new VB6FrameKind(rawProps, propertyGroupsFromParser, subControlsFromParser);
                        break;
                    case "CheckBox":
                        control.Kind = new VB6CheckBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "ComboBox":
                        control.Kind = new VB6ComboBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "CommandButton":
                        control.Kind = new VB6CommandButtonKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Data":
                        control.Kind = new VB6DataKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "DirListBox":
                        control.Kind = new VB6DirListBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "DriveListBox":
                        control.Kind = new VB6DriveListBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "FileListBox":
                        control.Kind = new VB6FileListBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Image":
                        control.Kind = new VB6ImageKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Label":
                        control.Kind = new VB6LabelKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Line":
                        control.Kind = new VB6LineKind(rawProps);
                        break;
                    case "ListBox":
                        control.Kind = new VB6ListBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "OLE":
                        // OLEKind constructor might need formFilePath and resourceResolver if it does its own specific FRX
                        // For now, assuming it's similar to other controls or handles FRX within its properties from rawProps.
                        // Let's assume for now the OLEProperties constructor and OleKind constructor handle it.
                        control.Kind = new VB6OleKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "OptionButton":
                        control.Kind = new VB6OptionButtonKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "PictureBox":
                        control.Kind = new VB6PictureBoxKind(rawProps, propertyGroupsFromParser, subControlsFromParser);
                        break;
                    case "HScrollBar":
                        control.Kind = new VB6HScrollBarKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "VScrollBar":
                        control.Kind = new VB6VScrollBarKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Shape":
                        control.Kind = new VB6ShapeKind(rawProps);
                        break;
                    case "TextBox":
                        control.Kind = new VB6TextBoxKind(rawProps, propertyGroupsFromParser);
                        break;
                    case "Timer":
                        control.Kind = new VB6TimerKind(rawProps);
                        break;
                    case "UserControl": // This is for the definition of a UserControl itself (from a .ctl file)
                        control.Kind = new VB6UserControlKind(
                            rawProps,
                            propertyGroupsFromParser,
                            subControlsFromParser, // Constituent controls *on* this UserControl
                            formFilePath, // Path to the .ctl file
                            resourceResolver
                        );
                        break;
                    case "ImageList":
                        // TODO
                        break;
                    default:
                        // Unhandled standard VB control, treat as custom.
                        //Console.Error.WriteLine($"Warning: BuildControl - Unhandled standard VB control kind '{fqn.Kind}' for '{fqn.Name}'. Treating as Custom.");
                        control.Kind = new VB6CustomKind(rawProps, propertyGroupsFromParser);
                        break;
                }
            }
            else
            {
                // Non-VB Namespace (e.g., RichTextLib.RichTextBox, MSComctlLib.ListViewCtrl, MSWinsockLib.Winsock)

                if (fqn.Namespace.Equals("RichTextLib", StringComparison.OrdinalIgnoreCase) && fqn.Kind.Equals("RichTextBox", StringComparison.OrdinalIgnoreCase))
                {
                    control.Kind = new VB6RichTextBoxKind(rawProps, propertyGroupsFromParser, formFilePath, resourceResolver);
                }
                else if (fqn.Namespace.Equals("MSWinsockLib", StringComparison.OrdinalIgnoreCase) && fqn.Kind.Equals("Winsock", StringComparison.OrdinalIgnoreCase))
                {
                    control.Kind = new VB6WinsockKind(rawProps, propertyGroupsFromParser);
                }
                else if (fqn.Namespace.Equals("MSComctlLib", StringComparison.OrdinalIgnoreCase) && (fqn.Kind.Equals("ListViewCtrl", StringComparison.OrdinalIgnoreCase) || fqn.Kind.Equals("ListView", StringComparison.OrdinalIgnoreCase))) // VB sometimes uses "ListView" as Kind
                {
                    var listViewKind = new VB6ListViewKind(rawProps, propertyGroupsFromParser, formFilePath, resourceResolver);
                    control.Kind = listViewKind;
                }
                else if (fqn.Namespace.Equals("MSComctlLib", StringComparison.OrdinalIgnoreCase) && fqn.Kind.Equals("TreeViewCtrl", StringComparison.OrdinalIgnoreCase))
                {
                    /* ... new VB6TreeViewKind(...) ... */
                }
                else if (fqn.Namespace.Equals("MSComDlg", StringComparison.OrdinalIgnoreCase) && fqn.Kind.Equals("CommonDialog", StringComparison.OrdinalIgnoreCase))
                {
                    /* ... new VB6CommonDialogKind(...) ... */
                }
                else // Fallback for any other (truly custom or unhandled) ActiveX control
                {
                    //Console.Error.WriteLine($"Warning: BuildControl - Unhandled ActiveX/Custom control '{fqn}'. Treating as Custom.");
                    control.Kind = new VB6CustomKind(rawProps, propertyGroupsFromParser);
                }
            }

            return control;
        }
        
        private static VB6MenuControl ConvertVB6ControlToMenuControl(VB6Control menuControlAsVB6Control)
        {
            if (menuControlAsVB6Control.Kind is VB6MenuKind actualMenuKind)
            {
                // Ensure TypedProperties is not null, which it shouldn't be if Kind constructor ran.
                MenuProperties typedProps = actualMenuKind.TypedProperties ?? new MenuProperties();

                return new VB6MenuControl
                {
                    Name = menuControlAsVB6Control.Name,
                    Tag = menuControlAsVB6Control.Tag,
                    Index = menuControlAsVB6Control.Index,
                    TypedProperties = typedProps, 
                    SubMenus = actualMenuKind.SubMenus, // These are already VB6MenuControl objects
                    PropertyGroups = actualMenuKind.PropertyGroups
                };
            }
            throw new InvalidOperationException($"Cannot convert VB6Control '{menuControlAsVB6Control.Name}' (Kind: {menuControlAsVB6Control.Kind?.GetType().Name}) to VB6MenuControl as its Kind is not VB6MenuKind.");
        }
    }
}