// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 OLE Container control.
    /// Used to embed or link OLE objects within a form.
    /// </summary>
    public class OleControl : Vb6ControlBase // VB6 Name: OLE
    {
        /// <summary>
        /// Gets or sets the specific properties for this OLE control.
        /// </summary>
        public OleProperties OleProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="OleControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "OLE1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The OLE-specific properties.</param>
        public OleControl(string name, int index, string tag, OleProperties properties)
            : base(name, index, tag)
        {
            OleProperties = properties ?? new OleProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OleControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public OleControl(string name, int index = -1)
            : this(name, index, string.Empty, new OleProperties())
        {
        }
    }
}