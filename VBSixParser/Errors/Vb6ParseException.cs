using System;
using System.Text;

namespace VB6Parse.Errors
{
    /// <summary>
    /// Represents errors that occur during VB6 parsing.
    /// </summary>
    public class Vb6ParseException : Exception
    {
        /// <summary>
        /// Gets the name of the file being parsed when the error occurred.
        /// Can be "unknown" if the file context is not available.
        /// </summary>
        public string FileName { get; }

        /// <summary>
        /// Gets the source code being parsed when the error occurred.
        /// Can be empty if the source code is not available.
        /// </summary>
        public string SourceCode { get; }

        /// <summary>
        /// Gets the zero-based offset within the source code where the error occurred.
        /// </summary>
        public int SourceOffset { get; }

        /// <summary>
        /// Gets the column number (one-based) where the error occurred.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the line number (one-based) where the error occurred.
        /// </summary>
        public int LineNumber { get; }

        /// <summary>
        /// Gets the specific kind of VB6 parsing error.
        /// </summary>
        public Vb6ErrorKind ErrorKind { get; }

        /// <summary>
        /// If the ErrorKind is 'Property', this specifies the more detailed property error.
        /// </summary>
        public PropertyErrorKind? SpecificPropertyError { get; }

        /// <summary>
        /// Gets the detailed error message.
        /// </summary>
        public override string Message => GetDetailedMessage();

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6ParseException"/> class
        /// with detailed information about the parsing error.
        /// </summary>
        /// <param name="fileName">The name of the file being parsed.</param>
        /// <param name="sourceCode">The source code being parsed.</param>
        /// <param name="sourceOffset">The zero-based offset in the source code.</param>
        /// <param name="column">The column number (one-based).</param>
        /// <param name="lineNumber">The line number (one-based).</param>
        /// <param name="kind">The kind of VB6 parsing error.</param>
        /// <param name="specificPropertyError">The specific property error kind, if applicable.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public Vb6ParseException(
            string fileName,
            string sourceCode,
            int sourceOffset,
            int column,
            int lineNumber,
            Vb6ErrorKind kind,
            PropertyErrorKind? specificPropertyError = null,
            Exception innerException = null)
            : base(FormatBaseMessage(fileName, lineNumber, column, kind, specificPropertyError), innerException)
        {
            FileName = fileName ?? "unknown";
            SourceCode = sourceCode ?? string.Empty;
            SourceOffset = sourceOffset;
            Column = column;
            LineNumber = lineNumber;
            ErrorKind = kind;
            SpecificPropertyError = specificPropertyError;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Vb6ParseException"/> class
        /// for errors where detailed stream information might not be available.
        /// </summary>
        /// <param name="kind">The kind of VB6 parsing error.</param>
        /// <param name="message">A custom message describing the error.</param>
        /// <param name="specificPropertyError">The specific property error kind, if applicable.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public Vb6ParseException(
            Vb6ErrorKind kind,
            string message = null,
            PropertyErrorKind? specificPropertyError = null,
            Exception innerException = null)
            : base(message ?? FormatBaseMessage("unknown", 0, 0, kind, specificPropertyError), innerException)
        {
            FileName = "unknown";
            SourceCode = string.Empty;
            SourceOffset = 0;
            Column = 0;
            LineNumber = 0;
            ErrorKind = kind;
            SpecificPropertyError = specificPropertyError;
        }
        
        // Standard constructors
        public Vb6ParseException() : base() { ErrorKind = Vb6ErrorKind.Unparseable; }
        public Vb6ParseException(string message) : base(message) { ErrorKind = Vb6ErrorKind.Unparseable; }
        public Vb6ParseException(string message, Exception innerException) : base(message, innerException) { ErrorKind = Vb6ErrorKind.Unparseable; }


        private static string FormatBaseMessage(string fileName, int lineNumber, int column, Vb6ErrorKind kind, PropertyErrorKind? specificKind)
        {
            var sb = new StringBuilder();
            sb.Append($"VB6 Parsing Error: {kind}");
            if (specificKind.HasValue)
            {
                sb.Append($": {specificKind.Value}");
            }
            if (fileName != "unknown" && fileName != null)
            {
                sb.Append($" in '{fileName}'");
            }
            if (lineNumber > 0 && column > 0)
            {
                sb.Append($" at Line {lineNumber}, Column {column}");
            }
            return sb.ToString();
        }

        private string GetDetailedMessage()
        {
            // This can be expanded to provide more context, similar to Rust's Display impl.
            // For now, it relies on the base message formatted in the constructor.
            // You could add snippets of SourceCode here if desired.
            return base.Message;
        }
    }
}