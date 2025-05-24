// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 ListBox control, used for displaying a list of items.
    /// It inherits common control properties and holds ListBox-specific properties.
    /// </summary>
    public class ListBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this ListBox control.
        /// </summary>
        public ListBoxProperties ListBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ListBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The ListBox-specific properties.</param>
        public ListBoxControl(string name, int index, string tag, ListBoxProperties properties)
            : base(name, index, tag)
        {
            ListBoxProperties = properties ?? new ListBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ListBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public ListBoxControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ListBoxProperties())
        {
        }
    }
}