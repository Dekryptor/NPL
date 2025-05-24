// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 TextBox control, inheriting common control properties
    /// and holding TextBox-specific properties.
    /// </summary>
    public class TextBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this TextBox control.
        /// </summary>
        public TextBoxProperties TextBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The TextBox-specific properties.</param>
        public TextBoxControl(string name, int index, string tag, TextBoxProperties properties)
            : base(name, index, tag)
        {
            TextBoxProperties = properties ?? new TextBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TextBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public TextBoxControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new TextBoxProperties())
        {
        }
    }
}