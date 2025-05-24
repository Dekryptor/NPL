// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Line control, used for drawing lines.
    /// It inherits common control properties and holds Line-specific properties.
    /// </summary>
    public class LineControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Line control.
        /// </summary>
        public LineProperties LineProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="LineControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (often system-generated like "Line1").</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The Line-specific properties.</param>
        public LineControl(string name, int index, string tag, LineProperties properties)
            : base(name, index, tag) // Line controls do have Name, Index, Tag
        {
            LineProperties = properties ?? new LineProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LineControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public LineControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new LineProperties())
        {
        }
    }
}