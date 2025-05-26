using System;
using System.Text;
using System.Linq; // For SequenceEqual on byte arrays
using VBSix.Errors; // For VB6ParseException, VB6ErrorKind, PropertyErrorType

namespace VBSix.Parsers
{
    /// <summary>
    /// Represents a checkpoint in a VB6Stream, allowing the stream
    /// to be reset to a previous state. This includes the byte index,
    /// line number, and column number.
    /// </summary>
    public readonly struct VB6StreamCheckpoint
    {
        public int Index { get; }
        public int LineNumber { get; }
        public int Column { get; }

        public VB6StreamCheckpoint(int index, int lineNumber, int column)
        {
            Index = index;
            LineNumber = lineNumber;
            Column = column;
        }

        public override string ToString()
        {
            return $"Index: {Index}, Line: {LineNumber}, Col: {Column}";
        }
    }

    /// <summary>
    /// A stream for parsing VB6 files, typically encoded in an ANSI codepage
    /// (like Windows-1252 for English systems).
    /// This stream manages the current position, line, and column numbers,
    /// and provides methods for peeking, advancing, and comparing byte sequences.
    /// It operates on a ReadOnlyMemory<byte> for efficient slicing.
    /// </summary>
    public class VB6Stream : ICloneable
    {
        // --- Static constructor to register encoding provider ---
        static VB6Stream()
        {
            // This ensures the provider is registered once when the VB6Stream type is first accessed.
            // This is crucial for .NET Core/.NET 5+ to support legacy encodings like Windows-1252.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        // --- End of static constructor ---

        public string FilePath { get; private set; }
        
        private readonly ReadOnlyMemory<byte> _fullStreamData; 
        
        public ReadOnlyMemory<byte> CurrentStream { get; private set; }

        /// <summary>
        /// The current 0-based byte index within the original _fullStreamData.
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// The current 1-based line number.
        /// </summary>
        public int LineNumber { get; private set; }

        /// <summary>
        /// The current 1-based column number on the current line.
        /// </summary>
        public int Column { get; private set; }

        /// <summary>
        /// Initializes a new instance of the VB6Stream class.
        /// </summary>
        /// <param name="filePath">The path of the file being parsed (for error reporting).</param>
        /// <param name="streamBytes">The byte content of the file.</param>
        public VB6Stream(string filePath, byte[] streamBytes)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            _fullStreamData = new ReadOnlyMemory<byte>(streamBytes ?? throw new ArgumentNullException(nameof(streamBytes)));
            CurrentStream = _fullStreamData;
            Index = 0;
            LineNumber = 1;
            Column = 1;
        }

        // Private constructor for cloning and advancing, maintaining reference to _fullStreamData
        private VB6Stream(string filePath, ReadOnlyMemory<byte> fullStreamData, ReadOnlyMemory<byte> currentStream, int index, int lineNumber, int column)
        {
            FilePath = filePath;
            _fullStreamData = fullStreamData;
            CurrentStream = currentStream;
            Index = index;
            LineNumber = lineNumber;
            Column = column;
        }

        public bool IsEmpty => CurrentStream.IsEmpty;
        public int Length => CurrentStream.Length;

        /// <summary>
        /// Advances the stream by a specified number of bytes.
        /// Updates Index, LineNumber, and Column accordingly.
        /// </summary>
        /// <param name="count">Number of bytes to advance. Must be non-negative and not exceed current stream length.</param>
        /// <returns>A new VB6Stream instance representing the stream after advancing.</returns>
        public VB6Stream Advance(int count)
        {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count), "Count cannot be negative.");
            if (count == 0) return new VB6Stream(FilePath, _fullStreamData, CurrentStream, Index, LineNumber, Column); // Return a new instance to maintain immutability pattern
            if (count > CurrentStream.Length)
                throw new ArgumentOutOfRangeException(nameof(count), $"Cannot advance by {count} bytes. Only {CurrentStream.Length} bytes remaining. File: {FilePath}, Line: {LineNumber}, Col: {Column}, Index: {Index}");

            int newIndex = Index;
            int newLineNumber = LineNumber;
            int newColumn = Column;

            ReadOnlySpan<byte> advancedBytes = CurrentStream.Span.Slice(0, count);
            for (int i = 0; i < count; i++)
            {
                newIndex++; // Overall index in _fullStreamData always increments
                byte currentByte = advancedBytes[i];

                if (currentByte == (byte)'\n')
                {
                    newLineNumber++;
                    newColumn = 1;
                }
                else if (currentByte == (byte)'\r')
                {
                    // If CR is followed by LF (within the `advancedBytes` slice), the LF will handle the newline.
                    // If CR is standalone, or at the very end of `advancedBytes` without an LF after it,
                    // then this CR itself constitutes a newline.
                    if (i + 1 < count && advancedBytes[i + 1] == (byte)'\n')
                    {
                        // This is part of a CRLF sequence. The '\n' will be processed next.
                        // Column still increments for the CR.
                        newColumn++;
                    }
                    else // Standalone CR or CR at end of consumed segment
                    {
                        newLineNumber++;
                        newColumn = 1;
                    }
                }
                else
                {
                    newColumn++;
                }
            }
            
            return new VB6Stream(
                this.FilePath,
                this._fullStreamData,
                this.CurrentStream.Slice(count),
                newIndex,
                newLineNumber,
                newColumn
            );
        }

        public VB6StreamCheckpoint Checkpoint() => new VB6StreamCheckpoint(Index, LineNumber, Column);

        public VB6Stream ResetToCheckpoint(VB6StreamCheckpoint checkpoint)
        {
            if (checkpoint.Index < 0 || checkpoint.Index > _fullStreamData.Length)
                throw new ArgumentOutOfRangeException(nameof(checkpoint), $"Invalid checkpoint index {checkpoint.Index}. Full stream length is {_fullStreamData.Length}.");

            return new VB6Stream(
                this.FilePath,
                this._fullStreamData,
                this._fullStreamData.Slice(checkpoint.Index), // CurrentStream becomes the slice from checkpoint's index to end
                checkpoint.Index,
                checkpoint.LineNumber,
                checkpoint.Column
            );
        }
        
        public int OffsetFrom(VB6StreamCheckpoint checkpoint) => Index - checkpoint.Index;
        
        public int FindByte(byte needle) => CurrentStream.Span.IndexOf(needle);

        public int FindSlice(string needle, bool ignoreCase = false)
        {
            if (string.IsNullOrEmpty(needle)) return -1;
            // VB6 files are typically Windows-1252 (superset of Latin1/ISO-8859-1) or system ANSI.
            // Encoding.Default attempts to use the system's current ANSI code page.
            // For consistency, explicitly using Windows-1252 (codepage 1252) is often safer for VB6.
            Encoding registeredEncoding = Encoding.GetEncoding(1252); 
            byte[] needleBytes = registeredEncoding.GetBytes(needle);

            if (!ignoreCase)
            {
                return CurrentStream.Span.IndexOf(needleBytes);
            }
            else
            {
                for (int i = 0; i <= CurrentStream.Length - needleBytes.Length; i++)
                {
                    bool match = true;
                    for (int j = 0; j < needleBytes.Length; j++)
                    {
                        // Case-insensitive comparison for single-byte characters
                        if (char.ToLowerInvariant((char)CurrentStream.Span[i + j]) != char.ToLowerInvariant((char)needleBytes[j]))
                        {
                            match = false;
                            break;
                        }
                    }
                    if (match) return i;
                }
                return -1;
            }
        }

        public bool CompareSlice(string expected)
        {
            if (string.IsNullOrEmpty(expected)) return CurrentStream.IsEmpty; 
            Encoding registeredEncoding = Encoding.GetEncoding(1252);
            byte[] expectedBytes = registeredEncoding.GetBytes(expected);
            if (CurrentStream.Length < expectedBytes.Length) return false;
            return CurrentStream.Span.Slice(0, expectedBytes.Length).SequenceEqual(expectedBytes);
        }

        public bool CompareSliceCaseless(string expected)
        {
            if (string.IsNullOrEmpty(expected)) return CurrentStream.IsEmpty;
            if (CurrentStream.Length < expected.Length) return false;
            for (int i = 0; i < expected.Length; i++)
            {
                // Using char.ToLowerInvariant for robust case-insensitivity with ASCII-compatible chars
                if (char.ToLowerInvariant((char)CurrentStream.Span[i]) != char.ToLowerInvariant(expected[i]))
                    return false;
            }
            return true;
        }

        public ReadOnlyMemory<byte> PeekSlice(int length)
        {
            if (length < 0) throw new ArgumentOutOfRangeException(nameof(length), "Length cannot be negative.");
            if (length > CurrentStream.Length)
                throw new ArgumentOutOfRangeException(nameof(length), $"Cannot peek {length} bytes, only {CurrentStream.Length} available. File: {FilePath}, Line: {LineNumber}, Col: {Column}, Index: {Index}");
            return CurrentStream.Slice(0, length);
        }
        
        public ReadOnlyMemory<byte> PeekSlice(int offset, int length)
        {
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset), "Offset cannot be negative.");
            if (length < 0) throw new ArgumentOutOfRangeException(nameof(length), "Length cannot be negative.");
            if (offset + length > CurrentStream.Length)
                throw new ArgumentOutOfRangeException(nameof(length), $"Cannot peek {length} bytes at offset {offset}, extends beyond current stream view length {CurrentStream.Length}. File: {FilePath}, Line: {LineNumber}, Col: {Column}, Index: {Index}");
            return CurrentStream.Slice(offset, length);
        }

        public byte? PeekByte() => CurrentStream.IsEmpty ? (byte?)null : CurrentStream.Span[0];
        
        public byte? PeekByteAt(int offset)
        {
             if (offset < 0 || offset >= CurrentStream.Length) return null;
             return CurrentStream.Span[offset];
        }

        private string GetSourceCodeSnippetForError(int contextChars = 30)
        {
            try
            {
                if (_fullStreamData.IsEmpty) return "[Empty Source]";

                int lineStart = Index;
                while (lineStart > 0 && _fullStreamData.Span[lineStart - 1] != (byte)'\n' && _fullStreamData.Span[lineStart - 1] != (byte)'\r')
                {
                    lineStart--;
                }

                int lineEnd = Index;
                while (lineEnd < _fullStreamData.Length && _fullStreamData.Span[lineEnd] != (byte)'\n' && _fullStreamData.Span[lineEnd] != (byte)'\r')
                {
                    lineEnd++;
                }

                // Use a specific encoding, now that the provider is registered.
                Encoding registeredEncoding = Encoding.GetEncoding(1252);
                string currentLine = registeredEncoding.GetString(_fullStreamData[lineStart..lineEnd].Span);
                string pointer = new string(' ', Column - 1) + "^";

                return $"{currentLine}{Environment.NewLine}{pointer}";
            }
            catch
            {
                int snippetStart = Math.Max(0, Index - contextChars);
                int snippetLength = Math.Min(contextChars * 2, _fullStreamData.Length - snippetStart);
                if (snippetLength <= 0) return "[End of Source]";
                // Use a specific encoding here too for the fallback
                Encoding registeredEncoding = Encoding.GetEncoding(1252);
                return registeredEncoding.GetString(_fullStreamData.Slice(snippetStart, snippetLength).Span) + (Index < _fullStreamData.Length - 1 ? "..." : "");
            }
        }

        public VB6ParseException CreateException(VB6ErrorKind kind, PropertyErrorType? propertyErrorDetail = null, Exception? innerException = null)
        {
            string snippet = GetSourceCodeSnippetForError();
             if (kind == VB6ErrorKind.Property)
            {
                 if (!propertyErrorDetail.HasValue) throw new ArgumentException("PropertyErrorDetail must be provided for Property kind error.", nameof(propertyErrorDetail));
                return VB6ParseException.FromPropertyError(propertyErrorDetail.Value, FilePath, snippet, Index, Column, LineNumber, innerException);
            }
            return new VB6ParseException(kind, FilePath, snippet, Index, Column, LineNumber, null, innerException);
        }

        public VB6ParseException CreateExceptionFromCheckpoint(VB6StreamCheckpoint checkpoint, VB6ErrorKind kind, PropertyErrorType? propertyErrorDetail = null, Exception? innerException = null)
        {
            string snippet;
            try
            {
                if (_fullStreamData.IsEmpty) snippet = "[Empty Source]";
                else
                {
                    int errorIndex = checkpoint.Index;
                    int lineStart = errorIndex;
                    while (lineStart > 0 && _fullStreamData.Span[lineStart - 1] != (byte)'\n' && _fullStreamData.Span[lineStart - 1] != (byte)'\r') lineStart--;
                    int lineEnd = errorIndex;
                    while (lineEnd < _fullStreamData.Length && _fullStreamData.Span[lineEnd] != (byte)'\n' && _fullStreamData.Span[lineEnd] != (byte)'\r') lineEnd++;
                    // Use a specific encoding here too for the fallback
                    Encoding registeredEncoding = Encoding.GetEncoding(1252);
                    string currentLine = registeredEncoding.GetString(_fullStreamData[lineStart..lineEnd].Span);
                    string pointer = new string(' ', checkpoint.Column - 1) + "^";
                    snippet = $"{currentLine}{Environment.NewLine}{pointer}";
                }
            }
            catch { snippet = "[Error retrieving source snippet from checkpoint]"; }

            if (kind == VB6ErrorKind.Property)
            {
                if (!propertyErrorDetail.HasValue) throw new ArgumentException("PropertyErrorDetail must be provided for Property kind error.", nameof(propertyErrorDetail));
                return VB6ParseException.FromPropertyError(propertyErrorDetail.Value, FilePath, snippet, checkpoint.Index, checkpoint.Column, checkpoint.LineNumber, innerException);
            }
            return new VB6ParseException(kind, FilePath, snippet, checkpoint.Index, checkpoint.Column, checkpoint.LineNumber, null, innerException);
        }

        public object Clone() => new VB6Stream(FilePath, _fullStreamData, CurrentStream, Index, LineNumber, Column);
    }
}