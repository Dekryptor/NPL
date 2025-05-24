// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Label control, inheriting common control properties
    /// and holding Label-specific properties.
    /// </summary>
    public class LabelControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Label control.
        /// </summary>
        public LabelProperties LabelProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="LabelControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The label-specific properties.</param>
        public LabelControl(string name, int index, string tag, LabelProperties properties)
            : base(name, index, tag)
        {
            LabelProperties = properties ?? new LabelProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LabelControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public LabelControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new LabelProperties())
        {
        }
    }
}