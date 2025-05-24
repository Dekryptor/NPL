// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 VScrollBar (Vertical ScrollBar) control.
    /// It inherits common control properties and uses ScrollBar-specific properties.
    /// </summary>
    public class VScrollBarControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this VScrollBar control.
        /// </summary>
        public ScrollBarProperties ScrollBarProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VScrollBarControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The ScrollBar-specific properties.</param>
        public VScrollBarControl(string name, int index, string tag, ScrollBarProperties properties)
            : base(name, index, tag)
        {
            ScrollBarProperties = properties ?? new ScrollBarProperties();
            // Optionally adjust default Width for VScrollBar if not already suitable
            if (properties == null) // Only if we are creating brand new default properties
            {
                 // Override the defaults from ScrollBarProperties which might be HScrollBar oriented
                this.ScrollBarProperties.Width = 285; // Typical Width for a VScrollBar
                this.ScrollBarProperties.Height = 1200; // Example Height for a VScrollBar
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VScrollBarControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public VScrollBarControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ScrollBarProperties()) // This will call the above constructor
        {
            // The default properties set in the primary constructor (when properties is null)
            // will establish VScrollBar specific dimensions.
        }
    }
}