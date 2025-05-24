// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Data control.
    /// Used for binding other controls to a database recordset.
    /// </summary>
    public class DataControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Data control.
        /// </summary>
        public DataProperties DataProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Data1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The Data-specific properties.</param>
        public DataControl(string name, int index, string tag, DataProperties properties)
            : base(name, index, tag)
        {
            DataProperties = properties ?? new DataProperties();
            // Set default caption if not already set from parsed properties
            if (string.IsNullOrEmpty(DataProperties.Caption) && !string.IsNullOrEmpty(name))
            {
                DataProperties.Caption = name;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public DataControl(string name, int index = -1)
            : this(name, index, string.Empty, new DataProperties())
        {
            // Ensure caption is set to name for default constructor too
            if (!string.IsNullOrEmpty(name))
            {
                 DataProperties.Caption = name;
            }
        }
    }
}