using System.Collections.Generic;
using VBSix.Parsers; // For PropertyValue
using VBSix.Utilities; // For PropertyParserHelpers if needed (though less common for truly custom controls)
using System.Linq; // For ToDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties of a custom or third-party ActiveX control
    /// for which no specific typed property class exists in this parsing library.
    /// It essentially holds a collection of all properties parsed from the .frm file
    /// for this control.
    /// </summary>
    /// <remarks>
    /// This class is used when a control's `Kind` in a .frm file (e.g., "MSComctlLib.ListViewCtrl")
    /// does not map to one of the standard VB6 controls explicitly handled by strongly-typed
    /// property classes (like `TextBoxProperties`, `CommandButtonProperties`, etc.).
    /// 
    /// The `RawProperties` dictionary will contain all key-value pairs as found in the
    /// .frm file for this custom control instance. Property values can be simple strings
    /// or binary data (if linked to an .frx file).
    /// 
    /// Property groups (`BeginProperty...EndProperty` blocks) for custom controls are
    /// stored in the `PropertyGroups` list within the `VB6CustomKind` class.
    /// </remarks>
    public class CustomControlProperties
    {
        /// <summary>
        /// A dictionary holding the raw property names and their corresponding values
        /// (text or binary resource) as parsed from the .frm file.
        /// Keys are case-insensitive, matching VB6 behavior.
        /// </summary>
        public Dictionary<string, PropertyValue> RawProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomControlProperties"/> class,
        /// typically with an empty property store.
        /// </summary>
        public CustomControlProperties()
        {
            RawProperties = new Dictionary<string, PropertyValue>(System.StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomControlProperties"/> class
        /// by copying properties from a provided raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values to initialize with.</param>
        public CustomControlProperties(IDictionary<string, PropertyValue> rawProps)
        {
            // Create a new dictionary, copying the entries. This ensures that this instance
            // has its own copy and is not just referencing the passed-in dictionary.
            RawProperties = rawProps != null
                ? new Dictionary<string, PropertyValue>(rawProps, System.StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, PropertyValue>(System.StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets the number of raw properties stored for this custom control.
        /// </summary>
        public int Count => RawProperties.Count;

        /// <summary>
        /// Gets a value indicating whether any raw properties are stored.
        /// </summary>
        public bool IsEmpty => RawProperties.Count == 0;

        // You can add helper methods here to access specific properties if there are common
        // patterns among custom controls you frequently encounter, e.g.:
        // public string? GetCustomStringProperty(string key, string? defaultValue = null) =>
        //     RawProperties.GetString(key, defaultValue ?? string.Empty); // Assuming extension method exists

        // public int GetCustomIntProperty(string key, int defaultValue = 0) =>
        //     RawProperties.GetInt32(key, defaultValue);
    }
}