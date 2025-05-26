using System;
using System.Collections.Generic;

namespace VBSix.Parsers
{
    /// <summary>
    /// Represents a VB6 file format version (e.g., "VERSION 1.0 CLASS").
    /// This is distinct from project or component versions.
    /// </summary>
    public class VB6FileFormatVersion
    {
        public byte Major { get; set; }
        public byte Minor { get; set; }

        public VB6FileFormatVersion(byte major, byte minor)
        {
            Major = major;
            Minor = minor;
        }

        public override string ToString()
        {
            return $"{Major}.{Minor:D2}"; // D2 ensures minor version is two digits, e.g., 5.00
        }
    }

    /// <summary>
    /// Specifies the kind of VB6 file header being parsed.
    /// This helps guide the version line parser, as some headers
    /// include a type keyword (like "CLASS") and others don't.
    /// </summary>
    public enum HeaderKind
    {
        /// <summary>
        /// Header for a VB6 Class Module (.cls file).
        /// Example: VERSION 1.0 CLASS
        /// </summary>
        Class,

        /// <summary>
        /// Header for a VB6 Form Module (.frm file).
        /// Example: VERSION 5.00
        /// </summary>
        Form,

        /// <summary>
        /// Header for a VB6 Standard Module (.bas file).
        /// Note: Standard modules typically don't have a "VERSION X.Y MODULE" line.
        /// Their header is usually just "Attribute VB_Name = ..."
        /// This enum member is included for completeness if a unified header parsing
        /// approach is attempted, but module parsing is often simpler.
        /// </summary>
        Module,
        
        UserControl, // .ctl files - VERSION 5.00 (similar to Forms)
        
		UserDocument, // .dob files - VERSION 5.00 (similar to Forms)
        
		PropertyPage // .pag files - VERSION 5.00 (similar to Forms)
        // These could be added if specific parsing rules for their version lines are needed.
        // For now, Form is a general category for files that just have "VERSION X.YY".
    }

    /// <summary>
    /// Represents if a class is in the global or local namespace.
    /// From `Attribute VB_GlobalNameSpace = True/False`.
    /// </summary>
    public enum NameSpaceKind
    {
        Global,
        /// <summary>The class is in the local namespace (default for new classes if attribute is False or absent).</summary>
        Local 
    }

    /// <summary>
    /// Determines if a class can be created externally.
    /// From `Attribute VB_Creatable = True/False`.
    /// </summary>
    public enum CreatableKind
    {
        False,
        /// <summary>The class can be created (VB6 default for a new public class is often effectively True if Exposed is True).</summary>
        True 
    }

    /// <summary>
    /// Determines if a class has a predeclared instance.
    /// From `Attribute VB_PredeclaredId = True/False`.
    /// </summary>
    public enum PreDeclaredIDKind
    {
        /// <summary>The class does not have a predeclared ID (default).</summary>
        False, 
        True
    }

    /// <summary>
    /// Determines if a class is publicly exposed or project-internal.
    /// From `Attribute VB_Exposed = True/False`.
    /// </summary>
    public enum ExposedKind
    {
		/// <summary>The class is not publicly exposed (project internal - default for new classes).</summary>
        False, // Default (Private / Project Internal)
		/// <summary>The class is publicly exposed.</summary>
        True   // Public
    }

    /// <summary>
    /// Represents the collection of attributes found in the header of
    /// VB6 files like .cls, .frm, .bas.
    /// Example:
    /// Attribute VB_Name = "MyClass"
    /// Attribute VB_GlobalNameSpace = False
    /// Attribute VB_Creatable = True
    /// Attribute VB_PredeclaredId = False
    /// Attribute VB_Exposed = True
    /// Attribute VB_Description = "A sample class"
    /// Attribute VB_Ext_Key = "MyCustomData", "SomeValue"
    /// </summary>
    public class VB6FileAttributes
    {
        /// <summary>
        /// The name of the component (e.g., class name, form name, module name).
        /// From `Attribute VB_Name = "..."`. This is a mandatory attribute for most VB6 files.
        /// </summary>
        public string Name { get; set; } = string.Empty;

        public NameSpaceKind GlobalNameSpace { get; set; } = NameSpaceKind.Local;
        
		// VB6 IDE default for a new Class Module:
        // Creatable = False, PredeclaredId = False, Exposed = False
		public CreatableKind Creatable { get; set; } = CreatableKind.True; // VB default for PublicNotCreatable implies True for Exposed, False for Creatable. Default Class module is False for all.
        public PreDeclaredIDKind PreDeclaredID { get; set; } = PreDeclaredIDKind.False;
        public ExposedKind Exposed { get; set; } = ExposedKind.False; // Default for new class is Private (Exposed=False, Creatable=False)
        
        /// <summary>
        /// Optional description string for the component.
        /// From `Attribute VB_Description = "..."`.
        /// </summary>
        public string? Description { get; set; }

        /// <summary>
        /// Stores any other `Attribute VB_Ext_KEY = "KeyString", "ValueString"` pairs,
        /// or other non-standard `Attribute Key = Value` lines.
        /// The dictionary key is the "KeyString" part from `VB_Ext_KEY` or the `Key` from other attributes.
        /// The dictionary value is the "ValueString" part or the `Value`.
        /// Keys are stored case-insensitively.
        /// </summary>
        public Dictionary<string, string> ExtKey { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public VB6FileAttributes() 
        {
            // Set defaults according to typical new VB6 items
            // A new Class Module in VB6 IDE defaults to:
            // VB_Name = "Class1" (set by parser)
            // VB_GlobalNameSpace = False
            // VB_Creatable = False
            // VB_PredeclaredId = False
            // VB_Exposed = False
        }
    }

    /// <summary>
    /// Helper class to represent a single parsed "Attribute Key = Value" line internally
    /// by the <see cref="HeaderParser"/> before its content is mapped to the strongly-typed
    /// fields of <see cref="VB6FileAttributes"/>.
    /// </summary>
    public class VB6Attribute
    {
		/// <summary>
        /// The key of the attribute (e.g., "VB_Name", "VB_GlobalNameSpace", "VB_Ext_KEY").
        /// </summary>
        public string Key { get; }
		
		/// <summary>
        /// The raw string value of the attribute as parsed from the file.
        /// For `VB_Ext_KEY`, this might be a composite string like `"KeyString", "ValueString"`.
        /// </summary>
        public string Value { get; }

        public VB6Attribute(string key, string value)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Value = value ?? string.Empty;
        }

        public override string ToString()
        {
            // For VB_Ext_KEY, Value might already include quotes and comma.
            // For others, Value is the direct content.
            // A more refined ToString might handle quoting based on content.
            if (Key.Equals("VB_Ext_KEY", StringComparison.OrdinalIgnoreCase) && !Value.Contains(","))
            {
                // This case is unlikely if parsed correctly, but for display:
                return $"Attribute {Key} = \"{Value}\""; // Assuming it should have been a single string value
            }
            return $"Attribute {Key} = {Value}"; 
        }
    }
}