// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 CommandButton control, inheriting common control properties
    /// and holding CommandButton-specific properties.
    /// </summary>
    public class CommandButtonControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this CommandButton control.
        /// </summary>
        public CommandButtonProperties CommandButtonProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandButtonControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        /// <param name="tag">The tag associated with the control.</param>
        /// <param name="properties">The command button-specific properties.</param>
        public CommandButtonControl(string name, int index, string tag, CommandButtonProperties properties)
            : base(name, index, tag)
        {
            CommandButtonProperties = properties ?? new CommandButtonProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandButtonControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public CommandButtonControl(string name, int index = -1) // Default index if not in an array
            : this(name, index, string.Empty, new CommandButtonProperties())
        {
        }
    }
}