// Note: The original prompt had namespace VBSix.Module.
// For consistency with other Parsers, using VBSix.Parsers.
// If a distinct "Module" model namespace is preferred, adjust accordingly.
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq; // For SequenceEqual
using VBSix.Errors;
using VBSix.Language; // For VB6Token

namespace VBSix.Parsers
{
    /// <summary>
    /// Represents a parsed VB6 Standard Module (.bas) file.
    /// A VB6 module file typically starts with an "Attribute VB_Name" line,
    /// followed by the module's source code.
    /// </summary>
    /// <example>
    /// <code>
    /// string moduleContent = @"Attribute VB_Name = ""Module1""
    /// Option Explicit
    /// 
    /// Public Sub MyProcedure()
    ///     Debug.Print ""Hello from MyProcedure""
    /// End Sub
    /// ";
    ///
    /// // Assuming Encoding.Default or appropriate encoding for VB6 files (e.g., Windows-1252)
    /// byte[] sourceBytes = Encoding.Default.GetBytes(moduleContent); 
    /// try
    /// {
    ///     VB6ModuleFile moduleFile = VB6ModuleFile.Parse("MyModule.bas", sourceBytes);
    ///     Console.WriteLine($"Module Name: {moduleFile.Name}");
    ///     Console.WriteLine($"Token Count: {moduleFile.Tokens.Count}");
    /// }
    /// catch (VB6ParseException ex)
    /// {
    ///     Console.Error.WriteLine($"Error parsing module: {ex.Message}");
    /// }
    /// </code>
    /// </example>
    public class VB6ModuleFile
    {
        /// <summary>
        /// The name of the module, extracted from the `Attribute VB_Name = "..."` line.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// A list of <see cref="VB6Token"/> objects representing the tokenized source code of the module.
        /// </summary>
        public List<VB6Token> Tokens { get; set; } = [];

        /// <summary>
        /// Initializes a new instance of the <see cref="VB6ModuleFile"/> class.
        /// Used for serialization or manual construction.
        /// </summary>
        public VB6ModuleFile()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VB6ModuleFile"/> class with a specified name and tokens.
        /// </summary>
        /// <param name="name">The name of the module.</param>
        /// <param name="tokens">The list of tokens representing the module's code.</param>
        /// <exception cref="ArgumentNullException">Thrown if name or tokens is null.</exception>
        public VB6ModuleFile(string name, List<VB6Token> tokens)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Tokens = tokens ?? throw new ArgumentNullException(nameof(tokens));
        }

        /// <summary>
        /// Parses the byte content of a VB6 standard module file (.bas).
        /// </summary>
        /// <param name="fileName">The name of the file being parsed (used for error reporting).</param>
        /// <param name="sourceCodeBytes">The byte array containing the module file content.</param>
        /// <returns>A <see cref="VB6ModuleFile"/> object representing the parsed module.</returns>
        /// <exception cref="ArgumentNullException">If fileName or sourceCodeBytes is null.</exception>
        /// <exception cref="ArgumentException">If fileName is empty or whitespace.</exception>
        /// <exception cref="VB6ParseException">If parsing fails due to invalid format or other errors.</exception>
        public static VB6ModuleFile Parse(string fileName, byte[] sourceCodeBytes)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("FileName cannot be empty.", nameof(fileName));
            ArgumentNullException.ThrowIfNull(sourceCodeBytes);

            VB6Stream stream = new(fileName, sourceCodeBytes);
            var initialStreamStateForError = stream.Checkpoint();

            try
            {
                VB6Stream s = HeaderParser.SkipSpace0(stream); // Skip leading whitespace if any
                
                // Standard modules *must* start with "Attribute VB_Name = "ModuleName""
                // Unlike class files, they don't have a "VERSION" line first.
                var (sAfterAttributeKeyword, _) = VB6.TakeSpecificString(s, "Attribute", caseSensitive: true);
                s = HeaderParser.SkipSpace1(sAfterAttributeKeyword);
                
                var (sAfterVBNameKeyword, _) = VB6.TakeSpecificString(s, "VB_Name", caseSensitive: true);
                s = HeaderParser.SkipSpace0(sAfterVBNameKeyword);
                
                s = VB6.ParseLiteral(s, (byte)'='); // Consume '='
                s = HeaderParser.SkipSpace0(s);

                string moduleName;
                var checkpointBeforeNameString = s.Checkpoint();
                if (s.IsEmpty || s.PeekByte() != (byte)'"')
                    throw s.CreateExceptionFromCheckpoint(checkpointBeforeNameString, VB6ErrorKind.StringParseError, innerException: new Exception("Module name not enclosed in quotes."));
                
                var (sAfterNameString, nameValue) = VB6.ParseVB6String(s); // Vb6.ParseVB6String consumes quotes and handles internal ""
                moduleName = nameValue;
                s = sAfterNameString;

                s = HeaderParser.ParseOptionalTrailingCommentAndLineEnding(s); // Consume rest of the attribute line
                
                List<VB6Token> tokens = VB6.Vb6Parse(s); // Parse the remaining stream as code tokens
                return new VB6ModuleFile(moduleName, tokens);
            }
            catch (VB6ParseException ex) 
            { 
                // If the error kind indicates something specific about the name attribute missing,
                // ensure it's correctly reported.
                if (ex.Kind == VB6ErrorKind.KeywordNotFound && ex.Message.Contains("Attribute") || 
                    ex.Kind == VB6ErrorKind.KeywordNotFound && ex.Message.Contains("VB_Name") ||
                    ex.Kind == VB6ErrorKind.NoEqualSplit)
                {
                     throw stream.CreateExceptionFromCheckpoint(initialStreamStateForError, VB6ErrorKind.MissingNameAttribute, innerException: ex);
                }
                throw; 
            }
            catch (Exception ex) // Catch-all for unexpected issues
            {
                throw stream.CreateExceptionFromCheckpoint(initialStreamStateForError, VB6ErrorKind.Unparseable, 
                    innerException: new Exception($"Unexpected error parsing module file '{fileName}'.", ex));
            }
        }
        
        public override bool Equals(object? obj) => 
            obj is VB6ModuleFile other && 
            Name == other.Name &&
            Tokens.SequenceEqual(other.Tokens); // Consider if a deeper token comparison is needed

        public override int GetHashCode() => HashCode.Combine(Name, Tokens.Count); // Simple hash based on name and token count

        public override string ToString() => $"VB6ModuleFile {{ Name = \"{Name}\", TokenCount = {Tokens.Count} }}";
    }
}