// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 DirListBox (Directory List Box) control.
    /// Used to display a hierarchical list of directories.
    /// </summary>
    public class DirListBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this DirListBox control.
        /// </summary>
        public DirListBoxProperties DirListBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DirListBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Dir1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The DirListBox-specific properties.</param>
        public DirListBoxControl(string name, int index, string tag, DirListBoxProperties properties)
            : base(name, index, tag)
        {
            DirListBoxProperties = properties ?? new DirListBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DirListBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public DirListBoxControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new DirListBoxProperties())
        {
        }
    }
}