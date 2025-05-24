// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 FileListBox control.
    /// Used to display a list of files in a specified directory, filterable by pattern and attributes.
    /// </summary>
    public class FileListBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this FileListBox control.
        /// </summary>
        public FileListBoxProperties FileListBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="FileListBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "File1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The FileListBox-specific properties.</param>
        public FileListBoxControl(string name, int index, string tag, FileListBoxProperties properties)
            : base(name, index, tag)
        {
            FileListBoxProperties = properties ?? new FileListBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FileListBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public FileListBoxControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new FileListBoxProperties())
        {
        }
    }
}