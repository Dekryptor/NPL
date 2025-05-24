// Namespace: VB6Parse.Parsers (or a more general Model namespace)
using System;
using System.Collections.Generic;

namespace VB6Parse.Parsers // Or VB6Parse.Model
{
    /// <summary>
    /// Represents a named group of properties, often found in VB6 forms
    /// for complex properties like Font, or for custom control property bags.
    /// Corresponds to the Rust VB6PropertyGroup struct.
    /// </summary>
    public class Vb6PropertyGroup
    {
        /// <summary>
        /// The name of the property group (e.g., "Font", "ImageListItems").
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// An optional GUID associated with the property group, sometimes used by ActiveX controls.
        /// </summary>
        public Guid? Guid { get; set; }

        /// <summary>
        /// A dictionary of properties within this group.
        /// The key is the property name.
        /// The value can be a simple string representation of the property's value,
        /// or another Vb6PropertyGroup instance for nested groups.
        /// </summary>
        public Dictionary<string, object> Properties { get; set; }

        public Vb6PropertyGroup(string name)
        {
            Name = name;
            Properties = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase); // VB property names are case-insensitive
        }
    }
}