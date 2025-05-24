// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 HScrollBar (Horizontal ScrollBar) control.
    /// It inherits common control properties and uses ScrollBar-specific properties.
    /// </summary>
    public class HScrollBarControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this HScrollBar control.
        /// </summary>
        public ScrollBarProperties ScrollBarProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="HScrollBarControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The ScrollBar-specific properties.</param>
        public HScrollBarControl(string name, int index, string tag, ScrollBarProperties properties)
            : base(name, index, tag)
        {
            ScrollBarProperties = properties ?? new ScrollBarProperties();
            // Optionally adjust default Height for HScrollBar if not already suitable
            // For example, if ScrollBarProperties has generic defaults:
            // if (properties == null) { this.ScrollBarProperties.Height = 285; /* Typical HScroll Height */ }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HScrollBarControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public HScrollBarControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ScrollBarProperties())
        {
            // Ensure HScrollBar specific default dimensions if ScrollBarProperties doesn't set them directionally
            // this.ScrollBarProperties.Height = 285; // A common default height for HScrollBars
            // this.ScrollBarProperties.Width = 1200; // A common default width
        }
    }
}