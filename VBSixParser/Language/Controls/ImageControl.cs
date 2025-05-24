// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Image control, used for displaying pictures.
    /// It inherits common control properties and holds Image-specific properties.
    /// </summary>
    public class ImageControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Image control.
        /// </summary>
        public ImageProperties ImageProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImageControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The Image-specific properties.</param>
        public ImageControl(string name, int index, string tag, ImageProperties properties)
            : base(name, index, tag)
        {
            ImageProperties = properties ?? new ImageProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImageControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public ImageControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new ImageProperties())
        {
        }
    }
}