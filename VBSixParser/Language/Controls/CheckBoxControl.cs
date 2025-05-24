// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 CheckBox control, inheriting common control properties
    /// and holding CheckBox-specific properties.
    /// </summary>
    public class CheckBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this CheckBox control.
        /// </summary>
        public CheckBoxProperties CheckBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CheckBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The CheckBox-specific properties.</param>
        public CheckBoxControl(string name, int index, string tag, CheckBoxProperties properties)
            : base(name, index, tag)
        {
            CheckBoxProperties = properties ?? new CheckBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CheckBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public CheckBoxControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new CheckBoxProperties())
        {
        }
    }
}