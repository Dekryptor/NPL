using System;

namespace VBSix.Parsers
{
    /// <summary>
    /// Represents an `Object=` line found in VB6 Project (.vbp) files or
    /// Form (.frm) files (for ActiveX controls). These lines define references
    /// to external components, which can be compiled libraries (identified by a GUID)
    /// or other VB6 sub-projects (identified by a path).
    /// </summary>
    /// <remarks>
    /// Examples from a .VBP file:
    /// Compiled Object: Object={00020430-0000-0000-C000-000000000046}#2.0#0; STDOLE2.TLB
    /// Sub-Project Object: Object=*\\AMySubProject.vbp
    ///
    /// Example from a .FRM file (ActiveX control):
    /// Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
    /// (Note: The FRM object line format can vary slightly, sometimes the filename is quoted).
    ///
    /// In C#, this is represented as a single class with nullable fields to accommodate both types.
    /// </remarks>
    public class VB6ObjectReference
    {
        /// <summary>
        /// The original raw string value as it appeared in the file after "Object=".
        /// Useful for debugging or if custom parsing of non-standard lines is needed.
        /// Example: "{GUID}#Version#Flags;FileName" or "*\\APath"
        /// </summary>
        public string OriginalValue { get; set; } = string.Empty;

        /// <summary>
        /// The GUID of the referenced object, if it's a compiled component.
        /// Stored as a string, typically including curly braces as found in the file,
        /// e.g., "{00020430-0000-0000-C000-000000000046}".
        /// This will be null if the reference is path-based (a sub-project).
        /// Can also be a raw GUID string without braces if parsed that way.
        /// </summary>
        public string? ObjectGuidString { get; set; } // Changed from ObjectGuid to clarify it's the string form
		
		/// <summary>
        /// The parsed <see cref="System.Guid"/> if <see cref="ObjectGuidString"/> represents a valid GUID.
        /// Null otherwise, or if this is a path-based reference.
        /// </summary>
        public Guid? ParsedGuid
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ObjectGuidString)) return null;
				
                string guidToParse = ObjectGuidString;
                if (guidToParse.StartsWith("{") && guidToParse.EndsWith("}"))
                {
                    guidToParse = guidToParse.Substring(1, guidToParse.Length - 2);
                }
				
                return Guid.TryParse(guidToParse, out Guid result) ? result : (Guid?)null;
            }
        }

        /// <summary>
        /// The version string of the object, if applicable (e.g., "2.0", "1.3").
        /// Present for compiled components.
        /// </summary>
        public string? Version { get; set; }

        /// <summary>
        /// An additional numeric or string value found in compiled object references.
        /// This often corresponds to flags or a Locale ID (LCID, typically "0" for neutral).
        /// Example: "0"
        /// </summary>
        public string? LocaleID { get; set; }

        /// <summary>
        /// The filename of the referenced component (e.g., "STDOLE2.TLB", "MSCOMCTL.OCX").
        /// Present for compiled components.
        /// </summary>
        public string? FileName { get; set; }

        /// <summary>
        /// The path to a referenced sub-project (e.g., "*\\ASubProj.vbp").
        /// This will be null if the reference is GUID-based.
        /// The path typically includes the "*\\A" or "*/A" prefix as stored in the VBP/FRM.
        /// </summary>
        public string? Path { get; set; }
		
		/// <summary>
        /// Indicates if this reference is to a sub-project (path-based)
        /// rather than a compiled component (GUID-based).
        /// Determined by whether the <see cref="Path"/> property is set.
        /// </summary>
        public bool IsSubProjectReference => !string.IsNullOrEmpty(Path);

        /// <summary>
        /// Indicates if this reference is to a compiled component (GUID-based).
        /// Determined by whether the <see cref="ObjectGuidString"/> property is set.
        /// </summary>
        public bool IsCompiledReference => !string.IsNullOrEmpty(ObjectGuidString);

        public override string ToString()
        {
            if (IsSubProjectReference)
            {
                return $"Project Object Ref: {Path}";
            }
            else if (IsCompiledReference)
            {
                return $"Compiled Object Ref: GUID={ObjectGuidString}, Version={Version}, File={FileName}";
            }
            return $"Object Ref (Unknown Type): {OriginalValue}";
        }
    }
}