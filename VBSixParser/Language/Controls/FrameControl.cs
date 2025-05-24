// Namespace: VB6Parse.Language.Controls
using System.Collections.Generic; // Required for List<Vb6ControlBase>

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Frame control. Frames are containers that can hold other controls
    /// and are used to group elements visually and functionally.
    /// </summary>
    public class FrameControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Frame control.
        /// </summary>
        public FrameProperties FrameProperties { get; set; }

        /// <summary>
        /// Gets or sets the list of controls contained within this Frame.
        /// In VB6 .frm files, controls placed inside a Frame are nested in its definition.
        /// </summary>
        public List<Vb6ControlBase> Controls { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="FrameControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Frame1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The Frame-specific properties.</param>
        public FrameControl(string name, int index, string tag, FrameProperties properties)
            : base(name, index, tag)
        {
            FrameProperties = properties ?? new FrameProperties();
            Controls = new List<Vb6ControlBase>(); // Initialize the list of contained controls.
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FrameControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public FrameControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new FrameProperties())
        {
        }
    }
}