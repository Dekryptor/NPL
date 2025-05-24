// Namespace: VB6Parse.Parsers (or VB6Parse.Model)
using System;

namespace VB6Parse.Model // Or VB6Parse.Parsers
{
    /// <summary>
    /// Represents an "Object=" line in a VB6 file, referencing an external
    /// component like an OCX or DLL.
    /// Example: Object={GUID}#VersionMajor.VersionMinor#Flags; FileName.ocx
    /// </summary>
    public class Vb6ObjectReference
    {
        /// <summary>
        /// The GUID of the referenced object.
        /// </summary>
        public Guid ObjectGuid { get; set; }

        /// <summary>
        /// The major version number of the object reference.
        /// </summary>
        public int VersionMajor { get; set; }

        /// <summary>
        /// The minor version number of the object reference.
        /// </summary>
        public int VersionMinor { get; set; }

        /// <summary>
        /// Licensing flags or other flags associated with the object reference.
        /// (The meaning of these flags, often 0, is not fully documented for all cases).
        /// </summary>
        public int Flags { get; set; } // Could be uint

        /// <summary>
        /// The filename of the component (e.g., "MSCOMCTL.OCX").
        /// This part appears after the semicolon.
        /// </summary>
        public string FileName { get; set; }
        
        /// <summary>
        /// The descriptive name or ProgID of the object, sometimes present.
        /// (e.g. "Microsoft Common Controls 6.0 (SP6)")
        /// This is often implied or part of the GUID registration rather than explicit in the Object= line itself,
        /// but the parser might resolve it or store the raw string.
        /// The Rust struct has `object_name` which might be the part before the semicolon, or a description.
        /// For now, we'll keep it simple. The raw string of the object line might also be useful.
        /// </summary>
        public string Description { get; set; } // Optional: Could be from a comment or registry.

        /// <summary>
        /// The full raw "Object=" line as it appeared in the file, excluding "Object=".
        /// Example: "{GUID}#Version#Flags; FileName.ocx"
        /// </summary>
        public string RawReferenceString { get; set; }


        public Vb6ObjectReference(string rawReference)
        {
            RawReferenceString = rawReference;
            FileName = string.Empty;
            Description = string.Empty;
            // Parsing logic for rawReference to populate other fields would go here or in the parser.
            // For now, we'll assume they are populated by the parser.
        }
         public Vb6ObjectReference()
        {
            RawReferenceString = string.Empty;
            FileName = string.Empty;
            Description = string.Empty;
        }

        public override string ToString()
        {
            return $"Object={{{ObjectGuid}}}#{VersionMajor}.{VersionMinor}#{Flags}; {FileName}";
        }
    }
}