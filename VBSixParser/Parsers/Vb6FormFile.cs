// Namespace: VB6Parse.Parsers (or VB6Parse.Model)
using System.Collections.Generic;
using VB6Parse.Language.Controls; // For Vb6Control

namespace VB6Parse.Model // Or VB6Parse.Parsers
{
    /// <summary>
    /// Represents a parsed VB6 Form file (.frm).
    /// This class structure is based on the Rust VB6FormFile struct.
    /// </summary>
    public class Vb6FormFile
    {
        /// <summary>
        /// The top-level Form control object, which contains all other controls and menus.
        /// </summary>
        public Vb6Control Form { get; set; }

        /// <summary>
        /// A list of "Object=" lines referencing external components (e.g., OCXs).
        /// </summary>
        public List<Vb6ObjectReference> ObjectReferences { get; set; }

        /// <summary>
        /// The file format version (e.g., "VERSION 5.00").
        /// </summary>
        public Vb6FileFormatVersion FormatVersion { get; set; }

        /// <summary>
        /// Global attributes of the form file (e.g., Attribute VB_Name = "Form1").
        /// </summary>
        public Vb6FileAttributes FileAttributes { get; set; }

        /// <summary>
        /// The text content of the VB6 code module embedded at the end of the .frm file.
        /// In the Rust version, this is stored as a list of VB6Tokens. For simplicity,
        /// we can start by storing the raw string or plan for tokenization later.
        /// </summary>
        public string CodeModuleText { get; set; }

        // Optional: Store the original file path
        public string FilePath { get; set; }


        public Vb6FormFile(string filePath = null)
        {
            FilePath = filePath;
            ObjectReferences = new List<Vb6ObjectReference>();
            FileAttributes = new Vb6FileAttributes();
            FormatVersion = new Vb6FileFormatVersion(); // Default or to be parsed
            CodeModuleText = string.Empty;
            // The 'Form' (Vb6Control) will be set by the parser.
        }
    }
}
