// Namespace: VB6Parse.Parsers (or VB6Parse.Model)
using System;
using System.Collections.Generic;

namespace VB6Parse.Model // Or VB6Parse.Parsers
{
    /// <summary>
    /// Represents the collection of "Attribute" lines in a VB6 file.
    /// Example: Attribute VB_Name = "Form1"
    /// </summary>
    public class Vb6FileAttributes
    {
        private readonly Dictionary<string, string> _attributes;

        public Vb6FileAttributes()
        {
            // VB attribute names are case-sensitive in some contexts but generally treated as case-insensitive by developers.
            // Storing them as read or canonicalizing (e.g. to upper) are options.
            // For simplicity, we'll use OrdinalIgnoreCase for retrieval, but store as read.
            _attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public void AddAttribute(string name, string value)
        {
            // Store the name as read, dictionary handles case-insensitivity for lookup.
            _attributes[name] = value;
        }

        public string GetAttribute(string name)
        {
            _attributes.TryGetValue(name, out string value);
            return value; // Returns null if not found
        }

        public IReadOnlyDictionary<string, string> GetAllAttributes()
        {
            return _attributes;
        }

        // Common known attributes as properties for convenience
        public string VbName
        {
            get => GetAttribute("VB_Name");
            set => AddAttribute("VB_Name", value);
        }

        public string VbGlobalNameSpace
        {
            get => GetAttribute("VB_GlobalNameSpace");
            set => AddAttribute("VB_GlobalNameSpace", value);
        }

        public string VbCreatable
        {
            get => GetAttribute("VB_Creatable");
            set => AddAttribute("VB_Creatable", value);
        }

        public string VbPredeclaredId
        {
            get => GetAttribute("VB_PredeclaredId");
            set => AddAttribute("VB_PredeclaredId", value);
        }

        public string VbExposed
        {
            get => GetAttribute("VB_Exposed");
            set => AddAttribute("VB_Exposed", value);
        }

        public string VbHelpID
        {
            get => GetAttribute("VB_HelpID");
            set => AddAttribute("VB_HelpID", value);
        }
        
        // Add other common attributes as needed...
    }
}