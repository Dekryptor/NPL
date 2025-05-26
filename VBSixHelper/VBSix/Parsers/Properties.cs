using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
// Assuming VB6Color, StartUpPosition will be in VBSix.Language
// Using directives for them will be needed in files that use GetColor/GetStartUpPosition helpers.
// The actual helper methods (GetString, GetInt32, GetEnum, etc.) are now in Utilities/PropertyParserHelpers.cs
// This Properties class is primarily a typed dictionary wrapper.

namespace VBSix.Parsers
{
    /// <summary>
    /// Represents a collection of properties for a VB6 control, typically parsed from a .frm file.
    /// Values can be simple text strings or binary data resolved from .frx files.
    /// Property keys are treated as case-insensitive, matching VB6 behavior.
    /// </summary>
    public class Properties
    {
        private readonly Dictionary<string, PropertyValue> _keyValueStore;

        /// <summary>
        /// Initializes a new instance of the <see cref="Properties"/> class.
        /// </summary>
        public Properties()
        {
            // VB6 property keys in .frm files are generally case-insensitive when accessed,
            // although they are stored with specific casing in the file.
            // StringComparer.OrdinalIgnoreCase ensures case-insensitive key lookups.
            _keyValueStore = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Inserts a property with a text value. If the key already exists, its value is updated.
        /// </summary>
        /// <param name="propertyKey">The key of the property.</param>
        /// <param name="textValue">The text value of the property.</param>
        /// <exception cref="ArgumentNullException">If propertyKey is null.</exception>
        public void Insert(string propertyKey, string textValue)
        {
            if (propertyKey == null) throw new ArgumentNullException(nameof(propertyKey));
            // textValue can be null or empty, which is valid for some properties.
            _keyValueStore[propertyKey] = new PropertyValue(textValue ?? string.Empty);
        }

        /// <summary>
        /// Inserts a property with a binary resource value. If the key already exists, its value is updated.
        /// The provided byte array is cloned to ensure the Properties collection owns its data.
        /// </summary>
        /// <param name="propertyKey">The key of the property.</param>
        /// <param name="resourceValue">The binary resource value of the property.</param>
        /// <exception cref="ArgumentNullException">If propertyKey or resourceValue is null.</exception>
        public void InsertResource(string propertyKey, byte[] resourceValue)
        {
            if (propertyKey == null) throw new ArgumentNullException(nameof(propertyKey));
            if (resourceValue == null) throw new ArgumentNullException(nameof(resourceValue));
            // PropertyValue constructor now handles cloning for resources
            _keyValueStore[propertyKey] = new PropertyValue(resourceValue);
        }

        /// <summary>
        /// Gets the number of properties in the collection.
        /// </summary>
        public int Count => _keyValueStore.Count;

        /// <summary>
        /// Gets a value indicating whether the collection is empty.
        /// </summary>
        public bool IsEmpty => _keyValueStore.Count == 0;

        /// <summary>
        /// Determines whether the collection contains a property with the specified key.
        /// </summary>
        /// <param name="propertyKey">The key to locate.</param>
        /// <returns>True if the key is found; otherwise, false.</returns>
        /// <exception cref="ArgumentNullException">If propertyKey is null.</exception>
        public bool ContainsKey(string propertyKey)
        {
            if (propertyKey == null) throw new ArgumentNullException(nameof(propertyKey));
            return _keyValueStore.ContainsKey(propertyKey);
        }

        /// <summary>
        /// Gets a collection containing the keys of the properties.
        /// </summary>
        /// <returns>An enumerable collection of keys.</returns>
        public IEnumerable<string> GetKeys() => _keyValueStore.Keys.ToList(); // ToList to return a snapshot

        /// <summary>
        /// Removes the property with the specified key from the collection.
        /// </summary>
        /// <param name="propertyKey">The key of the property to remove.</param>
        /// <returns>The <see cref="PropertyValue"/> that was removed, or null if the key was not found.</returns>
        /// <exception cref="ArgumentNullException">If propertyKey is null.</exception>
        public PropertyValue? Remove(string propertyKey)
        {
            if (propertyKey == null) throw new ArgumentNullException(nameof(propertyKey));
            if (_keyValueStore.TryGetValue(propertyKey, out var value))
            {
                _keyValueStore.Remove(propertyKey);
                return value;
            }
            return null;
        }

        /// <summary>
        /// Removes all properties from the collection.
        /// </summary>
        public void Clear() => _keyValueStore.Clear();

        /// <summary>
        /// Gets the <see cref="PropertyValue"/> associated with the specified key.
        /// </summary>
        /// <param name="propertyKey">The key of the property to get.</param>
        /// <returns>The <see cref="PropertyValue"/> associated with the key, or null if the key is not found.</returns>
        /// <exception cref="ArgumentNullException">If propertyKey is null.</exception>
        public PropertyValue? Get(string propertyKey)
        {
            if (propertyKey == null) throw new ArgumentNullException(nameof(propertyKey));
            _keyValueStore.TryGetValue(propertyKey, out var value);
            return value;
        }
        
        /// <summary>
        /// Provides direct access to the underlying dictionary store.
        /// Useful for extension methods or advanced manipulation.
        /// </summary>
        /// <returns>The internal dictionary storing property key-value pairs.</returns>
        public IDictionary<string, PropertyValue> GetRawStore() => _keyValueStore;

        // Helper methods like GetString, GetInt32, GetBoolean, GetVB6Color, GetEnum, etc.,
        // are now implemented as extension methods in Utilities/PropertyParserHelpers.cs
        // operating on IDictionary<string, PropertyValue>.
        // This keeps the Properties class focused on being a specialized dictionary.

        // Example of how to use the extension methods (assuming `props` is an instance of `Properties`):
        // string caption = props.GetRawStore().GetString("Caption", "DefaultCaption");
        // int height = props.GetRawStore().GetInt32("Height", 1000);
        // bool visible = props.GetRawStore().GetBoolean("Visible", true);
    }
}