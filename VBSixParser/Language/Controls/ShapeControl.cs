// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Shape control, used for drawing various graphical shapes.
    /// It inherits common control properties and holds Shape-specific properties.
    /// </summary>
    public class ShapeControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Shape control.
        /// </summary>
        public ShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (often system-generated like "Shape1").</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The Shape-specific properties.</param>
        public ShapeControl(string name, int index, string tag, ShapeProperties properties)
            : base(name, index, tag) // Shape controls do have Name, Index, Tag
        {
            ShapeProperties = properties ?? new ShapeProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public ShapeControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ShapeProperties())
        {
        }
    }
}