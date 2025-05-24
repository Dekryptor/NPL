// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 ComboBox control, which combines a text box with a list box.
    /// It inherits common control properties and holds ComboBox-specific properties.
    /// </summary>
    public class ComboBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this ComboBox control.
        /// </summary>
        public ComboBoxProperties ComboBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The ComboBox-specific properties.</param>
        public ComboBoxControl(string name, int index, string tag, ComboBoxProperties properties)
            : base(name, index, tag)
        {
            ComboBoxProperties = properties ?? new ComboBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComboBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public ComboBoxControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ComboBoxProperties())
        {
        }
    }
}