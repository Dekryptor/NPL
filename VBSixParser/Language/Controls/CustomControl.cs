// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a generic custom control (ActiveX / .OCX) in a VB6 form.
    /// This class is used when the control type is not one of the standard VB6 intrinsic controls.
    /// </summary>
    public class CustomControl : Vb6ControlBase
    {
        /// <summary>
        /// The GUID or ProgID that identifies the type of custom control (e.g., "MSComctlLib.TreeViewCtrl.2").
        /// This is crucial for understanding what kind of control this is.
        /// </summary>
        public string ObjectProgId { get; set; } // Or ObjectGUID as a Guid type

        /// <summary>
        /// Gets or sets the collection of custom properties for this control.
        /// These are properties specific to the OCX, stored as raw data.
        /// </summary>
        public CustomControlProperties CustomProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomControl"/> class.
        /// </summary>
        /// <param name="objectProgId">The ProgID or GUID of the custom control.</param>
        /// <param name="name">The name of the control instance.</param>
        /// <param name="index">The index if part of a control array.</param>
        /// <param name="tag">The tag string.</param>
        /// <param name="properties">The collection of custom properties.</param>
        public CustomControl(string objectProgId, string name, int index, string tag, CustomControlProperties properties)
            : base(name, index, tag)
        {
            ObjectProgId = objectProgId;
            CustomProperties = properties ?? new CustomControlProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomControl"/> class with default custom properties.
        /// </summary>
        /// <param name="objectProgId">The ProgID or GUID of the custom control.</param>
        /// <param name="name">The name of the control instance.</param>
        /// <param name="index">The index if part of a control array.</param>
        public CustomControl(string objectProgId, string name, int index = -1)
            : this(objectProgId, name, index, string.Empty, new CustomControlProperties())
        {
        }
    }
}