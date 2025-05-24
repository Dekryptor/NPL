using System;
using System.IO;
using System.Text; // For StringBuilder

// Namespace should match your project structure, e.g., VB6Parser.Parsers or VB6Parser.Utilities
namespace VB6Parser.Parsers // Or VB6Parser.Utilities
{
    public class Vb6StreamReader : IDisposable
    {
        private readonly StreamReader _reader;
        private string _pushedBackLine;
        public int LineNumber { get; private set; } // Tracks the starting line number of a logical line
        public int PhysicalLineNumber { get; private set; } // Tracks actual lines read from file
        public string FilePath { get; }

        public Vb6StreamReader(string filePath, Encoding encoding)
        {
            FilePath = filePath;
            _reader = new StreamReader(filePath, encoding);
            LineNumber = 0;
            PhysicalLineNumber = 0;
        }

        public string ReadLine()
        {
            if (_pushedBackLine != null)
            {
                string temp = _pushedBackLine;
                _pushedBackLine = null;
                // LineNumber for pushed back line is already set to its original starting line.
                return temp;
            }

            string currentSegment = _reader.ReadLine();
            if (currentSegment == null)
            {
                return null; // End of stream
            }

            PhysicalLineNumber++;
            LineNumber = PhysicalLineNumber; // Set the initial line number for this logical line

            // Handle line continuations
            StringBuilder logicalLineBuilder = new StringBuilder(currentSegment);
            
            // Trim trailing whitespace to accurately check for " _"
            string trimmedSegment = currentSegment.TrimEnd();

            while (trimmedSegment.EndsWith(" _"))
            {
                // Remove the " _" continuation marker (and any space before it if it was the only thing after content)
                // The TrimEnd() above handles spaces after '_'. We need to remove the " _" itself.
                // A simple approach: find the last occurrence of " _" on the trimmed line.
                int continuationMarkerPos = trimmedSegment.LastIndexOf(" _");
                if (continuationMarkerPos != -1) // Should always be true if EndsWith was true
                {
                    // Adjust builder length to remove " _" from the original (untrimmed) segment's representation
                    // This is tricky because currentSegment might have spaces after " _" that trimmedSegment doesn't.
                    // We want to remove from the *actual* " _" in currentSegment.
                    // Let's find the " _" in the original currentSegment based on the trimmed version.
                    // The length of `logicalLineBuilder` corresponds to `currentSegment`.
                    // `trimmedSegment.Length - 2` is where the content before " _" ends in the trimmed version.
                    // We need to map this back to `currentSegment`.
                    
                    // Simpler: just rebuild the builder up to the point of " _" from the original segment.
                    // This assumes " _" is the significant part.
                    // The part of currentSegment to keep is before the " _"
                    // This might be inaccurate if there are multiple " _" on one line (invalid VB but robust parsing might consider)
                    
                    // A more direct way:
                    // If currentSegment.TrimEnd().EndsWith(" _")
                    // then currentSegment looks like "blah blah _  " or "blah blah _"
                    // We need to remove the last " _" and any whitespace after it.
                    // The builder currently has `currentSegment`.
                    
                    // Find the last non-whitespace char before " _"
                    int actualContinuationPos = -1;
                    for(int i = currentSegment.Length -1; i >=0; i--)
                    {
                        if (currentSegment[i] == '_' && i > 0 && currentSegment[i-1] == ' ')
                        {
                            bool allSpaceAfter = true;
                            for(int j = i + 1; j < currentSegment.Length; j++)
                            {
                                if (!char.IsWhiteSpace(currentSegment[j]))
                                {
                                    allSpaceAfter = false;
                                    break;
                                }
                            }
                            if (allSpaceAfter)
                            {
                                actualContinuationPos = i - 1; // Position of the space before '_'
                                break;
                            }
                        }
                    }

                    if (actualContinuationPos != -1)
                    {
                        logicalLineBuilder.Length = actualContinuationPos; // Truncate to before " _"
                    }
                    else
                    {
                        // This case should ideally not be hit if TrimEnd().EndsWith(" _") was true.
                        // It implies " _" wasn't found as expected. Could be an issue if line is just " _".
                        // If line is " _", TrimEnd makes it " _", then EndsWith is true.
                        // logicalLineBuilder.Length should become 0.
                        if (trimmedSegment == " _") {
                            logicalLineBuilder.Length = 0;
                        } else {
                            // Fallback: if complex, just remove last 2 chars from trimmed and hope for best.
                            // This part needs to be robust.
                            // A simpler assumption: if currentSegment.TrimEnd().EndsWith(" _"), then
                            // the content part is currentSegment up to the space of the *last* " _".
                            // This is what the Rust parser `line_continuation` would typically handle more gracefully.
                            // For now, let's use the `actualContinuationPos` logic.
                        }
                    }


                } // else something is odd, " _" was not found as expected.

                string nextPhysicalLine = _reader.ReadLine();
                if (nextPhysicalLine == null)
                {
                    break; // End of stream, treat current accumulated line as complete
                }
                PhysicalLineNumber++;
                // Do NOT update LineNumber here, it stays as the start of the logical line.
                
                logicalLineBuilder.Append(nextPhysicalLine.TrimStart()); // VB often has indentation on continued lines
                currentSegment = nextPhysicalLine; // For the next loop iteration's EndsWith check
                trimmedSegment = currentSegment.TrimEnd();
            }

            return logicalLineBuilder.ToString();
        }

        public void PushBack(string line)
        {
            if (_pushedBackLine != null)
            {
                // Consider if multiple pushbacks are needed or if this is an error.
                // For simplicity, matching previous behavior:
                throw new InvalidOperationException("Cannot push back more than one line.");
            }
            _pushedBackLine = line;
            // When pushing back, LineNumber is already the start of this (potentially multi-physical) logical line.
            // We don't need to decrement LineNumber here.
            // PhysicalLineNumber is not decremented as it reflects the actual read position.
        }
        
        public bool EndOfStream => _reader.EndOfStream && _pushedBackLine == null;

        public void Close() => _reader.Close();
        public void Dispose() => _reader.Dispose();
    }
}