// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 DriveListBox control.
    /// Displays a list of available disk drives in the system.
    /// </summary>
    public class DriveListBoxControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this DriveListBox control.
        /// </summary>
        public DriveListBoxProperties DriveListBoxProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DriveListBoxControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Drive1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The DriveListBox-specific properties.</param>
        public DriveListBoxControl(string name, int index, string tag, DriveListBoxProperties properties)
            : base(name, index, tag)
        {
            DriveListBoxProperties = properties ?? new DriveListBoxProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DriveListBoxControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public DriveListBoxControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new DriveListBoxProperties())
        {
        }
    }
}