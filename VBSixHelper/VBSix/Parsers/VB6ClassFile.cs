using VBSix.Language;
using VBSix.Errors;

namespace VBSix.Parsers
{
    // --- Enums specific to Class File properties ---

    /// <summary>
    /// Specifies the instancing behavior for a class when used as a COM object.
    /// Corresponds to the `MultiUse` property in a .cls file header.
    /// </summary>
    public enum FileUsage
    {
        /// <summary>
        /// A single instance of the class object will be created for all clients.
        /// VBP Value: -1 (True). This is often the default for public creatable classes.
        /// </summary>
        MultiUse = -1,
        /// <summary>
        /// A new instance of the class object will be created for each client.
        /// VBP Value: 0 (False).
        /// </summary>
        SingleUse = 0
    }

    /// <summary>
    /// Specifies whether instances of a class can be saved to disk (persisted).
    /// Corresponds to the `Persistable` property in a .cls file header.
    /// </summary>
    public enum Persistance
    {
        /// <summary>
        /// The class cannot be saved to a property bag. VBP Value: 0 (False). (Default)
        /// </summary>
        NonPersistable = 0,
        /// <summary>
        /// The class can be saved to a property bag. VBP Value: -1 (True).
        /// Adds InitProperties, ReadProperties, WriteProperties events, and PropertyChanged method.
        /// </summary>
        Persistable = -1
    }

    /// <summary>
    /// Specifies the Microsoft Transaction Server (MTS) transaction mode for a class.
    /// Corresponds to the `MTSTransactionMode` property in a .cls file header.
    /// </summary>
    public enum MtsStatus
    {
        /// <summary>Not an MTS component. VBP Value: 0. (Default)</summary>
        NotAnMTSObject = 0,
        /// <summary>MTS component, does not support transactions. VBP Value: 1.</summary>
        NoTransactions = 1,
        /// <summary>MTS component, requires a transaction. VBP Value: 2.</summary>
        RequiresTransaction = 2,
        /// <summary>MTS component, uses an existing transaction. VBP Value: 3.</summary>
        UsesTransaction = 3,
        /// <summary>MTS component, requires a new transaction. VBP Value: 4.</summary>
        RequiresNewTransaction = 4
    }

    /// <summary>
    /// Specifies whether a class can act as a data source.
    /// Corresponds to the `DataSourceBehavior` property in a .cls file header.
    /// </summary>
    public enum DataSourceBehavior
    {
        /// <summary>Class does not support acting as a Data Source. VBP Value: 0. (Default)</summary>
        None = 0,
        /// <summary>Class supports acting as a Data Source. VBP Value: 1.</summary>
        DataSource = 1
    }

    /// <summary>
    /// Specifies the data binding behavior for a class.
    /// Corresponds to the `DataBindingBehavior` property in a .cls file header.
    /// </summary>
    public enum DataBindingBehavior
    {
        /// <summary>Class does not support data binding. VBP Value: 0. (Default)</summary>
        None = 0,
        /// <summary>Class supports simple data binding. VBP Value: 1.</summary>
        Simple = 1,
        /// <summary>Class supports complex data binding. VBP Value: 2.</summary>
        Complex = 2
    }

    // --- Class Property and Header Structures ---

    /// <summary>
    /// Holds the properties found in the `BEGIN`...`END` block of a VB6 class file header.
    /// These properties define COM behavior, persistence, and data binding capabilities.
    /// </summary>
    public class VB6ClassProperties
    {
        public FileUsage MultiUse { get; set; } = FileUsage.MultiUse;
        public Persistance Persistable { get; set; } = Persistance.NonPersistable;
        public DataBindingBehavior DataBindingBehavior { get; set; } = DataBindingBehavior.None;
        public DataSourceBehavior DataSourceBehavior { get; set; } = DataSourceBehavior.None;
        public MtsStatus MtsTransactionMode { get; set; } = MtsStatus.NotAnMTSObject;
    }

    /// <summary>
    /// Represents the complete header of a VB6 class file, including version,
    /// `BEGIN`...`END` block properties, and `Attribute` statements.
    /// </summary>
    public class VB6ClassHeader
    {
        public VB6FileFormatVersion Version { get; set; }
        public VB6ClassProperties Properties { get; set; }
        public VB6FileAttributes Attributes { get; set; }

        public VB6ClassHeader(VB6FileFormatVersion version, VB6ClassProperties properties, VB6FileAttributes attributes)
        {
            Version = version ?? throw new ArgumentNullException(nameof(version));
            Properties = properties ?? throw new ArgumentNullException(nameof(properties));
            Attributes = attributes ?? throw new ArgumentNullException(nameof(attributes));
        }
    }

    /// <summary>
    /// Represents a parsed VB6 Class Module (.cls) file.
    /// This includes its header information (version, attributes, COM properties)
    /// and the tokenized source code body.
    /// </summary>
    /// <example>
    /// <code>
    /// string classContent = @"VERSION 1.0 CLASS
    /// BEGIN
    ///   MultiUse = -1  'True
    ///   Persistable = 0  'NotPersistable
    ///   DataBindingBehavior = 0  'vbNone
    ///   DataSourceBehavior = 0  'vbNone
    ///   MTSTransactionMode = 0  'NotAnMTSObject
    /// END
    /// Attribute VB_Name = ""Something""
    /// Attribute VB_GlobalNameSpace = False
    /// Attribute VB_Creatable = True
    /// Attribute VB_PredeclaredId = False
    /// Attribute VB_Exposed = False
    /// 
    /// Public Sub MyMethod()
    ///     MsgBox ""Hello from MyMethod""
    /// End Sub
    /// ";
    ///
    /// byte[] sourceBytes = Encoding.Default.GetBytes(classContent); // Or appropriate encoding
    /// try
    /// {
    ///     VB6ClassFile classFile = VB6ClassFile.Parse("MyClass.cls", sourceBytes);
    ///     Console.WriteLine($"Class Name: {classFile.Header.Attributes.Name}");
    ///     Console.WriteLine($"File Format Version: {classFile.Header.Version}");
    ///     Console.WriteLine($"MultiUse: {classFile.Header.Properties.MultiUse}");
    ///     Console.WriteLine($"Token Count: {classFile.Tokens.Count}");
    /// }
    /// catch (VB6ParseException ex)
    /// {
    ///     Console.Error.WriteLine($"Error parsing class: {ex.Message}");
    /// }
    /// </code>
    /// </example>
    public class VB6ClassFile
    {
        public VB6ClassHeader Header { get; }
        public List<VB6Token> Tokens { get; }

        public VB6ClassFile(VB6ClassHeader header, List<VB6Token> tokens)
        {
            Header = header ?? throw new ArgumentNullException(nameof(header));
            Tokens = tokens ?? [];
        }

        /// <summary>
        /// Parses the byte content of a VB6 class file.
        /// </summary>
        /// <param name="fileName">The name of the class file (for error reporting).</param>
        /// <param name="sourceCodeBytes">The byte array containing the class file content.</param>
        /// <returns>A <see cref="VB6ClassFile"/> object representing the parsed class.</returns>
        /// <exception cref="ArgumentNullException">If fileName or sourceCodeBytes is null.</exception>
        /// <exception cref="ArgumentException">If fileName is empty or whitespace.</exception>
        /// <exception cref="VB6ParseException">If parsing fails due to invalid format or other errors.</exception>
        public static VB6ClassFile Parse(string fileName, byte[] sourceCodeBytes)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("FileName cannot be empty.", nameof(fileName));
            ArgumentNullException.ThrowIfNull(sourceCodeBytes);

            VB6Stream stream = new(fileName, sourceCodeBytes);
            var initialStreamForError = stream.Checkpoint(); // For top-level errors

            try
            {
                var (header, streamAfterHeader) = ParseClassHeader(stream);
                List<VB6Token> tokens = VB6.Vb6Parse(streamAfterHeader);
                return new VB6ClassFile(header, tokens);
            }
            catch (VB6ParseException) 
            { 
                throw; // Re-throw known parser errors
            }
            catch (Exception ex) // Catch-all for unexpected issues
            {
                throw stream.CreateExceptionFromCheckpoint(initialStreamForError, VB6ErrorKind.Unparseable, 
                    innerException: new Exception($"Unexpected error parsing class file '{fileName}'.", ex));
            }
        }

        private static (VB6ClassHeader header, VB6Stream nextStream) ParseClassHeader(VB6Stream stream)
        {
            var initialHeaderCheckpoint = stream.Checkpoint();
            try
            {
                var (sAfterVersion, version, actualKind) = HeaderParser.ParseVersionLine(stream, HeaderKind.Class);
                if (actualKind != HeaderKind.Class) 
                    throw stream.CreateExceptionFromCheckpoint(initialHeaderCheckpoint,VB6ErrorKind.Header, innerException: new Exception($"Expected CLASS header kind, but found '{actualKind}'."));
                
                var (properties, sAfterProps) = ParseClassPropertiesBlock(sAfterVersion);
                var (sAfterAttrs, attributes) = HeaderParser.ParseAttributes(sAfterProps);

                return (new VB6ClassHeader(version, properties, attributes), sAfterAttrs);
            }
            catch (VB6ParseException) { throw; } // Propagate specific parse errors
            catch (Exception ex)
            {
                throw stream.CreateExceptionFromCheckpoint(initialHeaderCheckpoint, VB6ErrorKind.Header, innerException: new Exception("Generic error parsing class header content.", ex));
            }
        }

        private static (VB6ClassProperties properties, VB6Stream nextStream) ParseClassPropertiesBlock(VB6Stream stream)
        {
            var initialBlockCheckpoint = stream.Checkpoint();
            VB6Stream s = HeaderParser.SkipSpace0(stream); 

            // Parse BEGIN
            var (sAfterBeginKeyword, _) = VB6.TakeSpecificString(s, "BEGIN", caseSensitive: true); // Throws if not found
            VB6Stream sAfterBeginLine = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(sAfterBeginKeyword);
            s = sAfterBeginLine;

            var props = new VB6ClassProperties();
            bool endFound = false;

            while(!s.IsEmpty)
            {
                var lineCheckpoint = s.Checkpoint();
                VB6Stream tempStreamForEndCheck = HeaderParser.SkipSpace0(s); // Use a temporary stream to check for END without consuming non-END line's spaces

                if (tempStreamForEndCheck.CompareSliceCaseless("END"))
                {
                    // If "END" is found, consume it from the original stream 's' (which might have leading spaces for the END line)
                    s = HeaderParser.SkipSpace0(s); // Ensure 's' is at "END"
                    var(sAfterEndKeyword, _) = VB6.TakeSpecificString(s, "END", caseSensitive: true);
                    s = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(sAfterEndKeyword);
                    endFound = true;
                    break;
                }
                
                // If not END, parse key = value from the current 's'
                var (sAfterKeyValue, key, value) = HeaderParser.ParseKeyValueLine(s); // ParseKeyValueLine consumes its own EOL
                s = sAfterKeyValue;

                switch (key) // Case-insensitive by default if Dictionary uses StringComparer.OrdinalIgnoreCase
                {
                    case "Persistable":
                        if (value == "-1") props.Persistable = Persistance.Persistable;
                        else if (value == "0") props.Persistable = Persistance.NonPersistable;
                        else throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.InvalidPropertyValueZeroNegOne, new Exception($"Invalid value '{value}' for Persistable."));
                        break;
                    case "MultiUse":
                        if (value == "-1") props.MultiUse = FileUsage.MultiUse;
                        else if (value == "0") props.MultiUse = FileUsage.SingleUse;
                        else throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.InvalidPropertyValueZeroNegOne, new Exception($"Invalid value '{value}' for MultiUse."));
                        break;
                    case "DataBindingBehavior":
                        if (value == "0") props.DataBindingBehavior = DataBindingBehavior.None;
                        else if (value == "1") props.DataBindingBehavior = DataBindingBehavior.Simple;
                        else if (value == "2") props.DataBindingBehavior = DataBindingBehavior.Complex;
                        else throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.UnknownProperty, new Exception($"Invalid value '{value}' for DataBindingBehavior."));
                        break;
                    case "DataSourceBehavior":
                        if (value == "0") props.DataSourceBehavior = DataSourceBehavior.None;
                        else if (value == "1") props.DataSourceBehavior = DataSourceBehavior.DataSource;
                        else throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.UnknownProperty, new Exception($"Invalid value '{value}' for DataSourceBehavior."));
                        break;
                    case "MTSTransactionMode":
                        if (!Enum.TryParse<MtsStatus>(value, out var mtsStatus) || !Enum.IsDefined(mtsStatus)) // Check if valid enum member
                             throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.UnknownProperty, new Exception($"Invalid value '{value}' for MTSTransactionMode."));
                        props.MtsTransactionMode = mtsStatus;
                        break;
                    default:
                         throw s.CreateExceptionFromCheckpoint(lineCheckpoint, VB6ErrorKind.Property, PropertyErrorType.UnknownProperty, new Exception($"Unknown class property key: '{key}'."));
                }
            }

            if (!endFound)
                throw stream.CreateExceptionFromCheckpoint(initialBlockCheckpoint, VB6ErrorKind.NoEndKeyword, innerException: new Exception("Missing END statement for class properties block."));

            return (props, s);
        }
    }
}