// Namespace: VB6Parse.Language.Controls

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Represents a VB6 Timer control.
    /// This control is invisible at runtime and is used to trigger Timer events at set intervals.
    /// </summary>
    public class TimerControl : Vb6ControlBase
    {
        /// <summary>
        /// Gets or sets the specific properties for this Timer control.
        /// </summary>
        public TimerProperties TimerProperties { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimerControl"/> class.
        /// </summary>
        /// <param name="name">The name of the control (e.g., "Timer1").</param>
        /// <param name="index">The index of the control if it's part of a control array.</param>
        /// <param name="tag">The tag string associated with the control.</param>
        /// <param name="properties">The Timer-specific properties.</param>
        public TimerControl(string name, int index, string tag, TimerProperties properties)
            : base(name, index, tag)
        {
            TimerProperties = properties ?? new TimerProperties();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="TimerControl"/> class with default properties.
        /// </summary>
        /// <param name="name">The name of the control.</param>
        /// <param name="index">The index of the control (if part of an array).</param>
        public TimerControl(string name, int index = -1) // Default index if not in an array.
            : this(name, index, string.Empty, new TimerProperties())
        {
        }
    }
}