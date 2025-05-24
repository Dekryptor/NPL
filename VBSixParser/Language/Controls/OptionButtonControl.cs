// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 OptionButton control (Radio Button), inheriting common control properties
    /// and holding OptionButton-specific properties.
    /// </summary>
    public class OptionButtonControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this OptionButton control.
        /// </summary>
        public OptionButtonProperties OptionButtonProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="OptionButtonControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The OptionButton-specific properties.</param>
        public OptionButtonControl(string name, int index, string tag, OptionButtonProperties properties)
            : base(name, index, tag)
        {
            OptionButtonProperties = properties ?? new OptionButtonProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OptionButtonControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public OptionButtonControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new OptionButtonProperties())
        {
        }
    }
}