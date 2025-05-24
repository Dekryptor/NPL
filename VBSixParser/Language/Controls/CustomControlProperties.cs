// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic;

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents the properties of a custom control (e.g., an ActiveX control / .OCX).
    /// Properties are stored in a general-purpose collection as their specific types
    /// may not be known in advance by the parser.
    /// </summary>
    public class CustomControlProperties
    {
        /// <summary>
        /// A dictionary storing the properties of the custom control.
        /// Keys are property names, and values are their raw byte representations
        /// as found in the form file.
        /// </summary>
        /// <remarks>
        /// Standard properties like Name, Index, Tag, Left, Top, Width, Height,
        /// Visible, Enabled, etc., might be parsed into the parent Vb6ControlBase
        /// or a more specific custom control base, while this collection holds
        /// properties unique to the specific OCX.
        /// </remarks>
        public Dictionary<string, byte[]> PropertyStore { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomControlProperties"/> class.
        /// </summary>
        public CustomControlProperties()
        {
            PropertyStore = new Dictionary<string, byte[]>(System.StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets the number of custom properties stored.
        /// </summary>
        public int Count => PropertyStore.Count;

        /// <summary>
        /// Adds a property to the store.
        /// </summary>
        /// <param name="name">The name of the property.</param>
        /// <param name="rawValue">The raw byte value of the property.</param>
        public void AddProperty(string name, byte[] rawValue)
        {
            PropertyStore[name] = rawValue;
        }

        /// <summary>
        /// Tries to get a property's raw value.
        /// </summary>
        /// <param name="name">The name of the property.</param>
        /// <param name="rawValue">The raw byte value of the property, if found.</param>
        /// <returns>True if the property was found, otherwise false.</returns>
        public bool TryGetProperty(string name, out byte[] rawValue)
        {
            return PropertyStore.TryGetValue(name, out rawValue);
        }
    }
}