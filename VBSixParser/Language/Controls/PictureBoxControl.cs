// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic; // Required for List<Vb6ControlBase>

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 PictureBox control. It can display images, act as a container
    /// for other controls, and provide a surface for graphical operations.
    /// </summary>
    public class PictureBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this PictureBox control.
        /// </summary>
        public PictureBoxProperties PictureBoxProperties { get; set; }

        /// <summary>
        /// Gets or sets the list of controls contained within this PictureBox.
        /// </summary>
        public List<Vb6ControlBase> Controls { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Picture1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The PictureBox-specific properties.</param>
        public PictureBoxControl(string name, int index, string tag, PictureBoxProperties properties)
            : base(name, index, tag)
        {
            PictureBoxProperties = properties ?? new PictureBoxProperties();
            Controls = new List<Vb6ControlBase>(); // Initialize the list of contained controls.
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PictureBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public PictureBoxControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new PictureBoxProperties())
        {
        }
    }
}