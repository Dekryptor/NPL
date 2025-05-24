// Namespace: VB6Parse.Language.Controls
// For common enums like Activation.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Timer control.
    /// The Timer control is invisible at runtime and triggers an event at specified intervals.
    /// </summary>
    public class TimerProperties
    {
        /// <summary>
        /// Gets or sets a value indicating whether the Timer control can respond to user interaction.
        /// At runtime, if Enabled is True, the timer will fire events. If False, it will not.
        /// </summary>
        public Activation Enabled { get; set; }

        /// <summary>
        /// Gets or sets the interval, in milliseconds, between Timer events.
        /// A value of 0 disables the timer. The maximum value is 65,535.
        /// </summary>
        public int Interval { get; set; }

        /// <summary>
        /// Gets or sets the design-time left position of the Timer icon on the form.
        /// This property has no effect at runtime as the Timer is invisible.
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Gets or sets the design-time top position of the Timer icon on the form.
        /// This property has no effect at runtime as the Timer is invisible.
        /// </summary>
        public int Top { get; set; }


        /// <summary>
        /// Initializes a new instance of the <see cref="TimerProperties"/> class with default VB6 values.
        /// </summary>
        public TimerProperties()
        {
            // VB6 defaults for Timer control:
            Enabled = Activation.Enabled; // VB6 default is True (Enabled). Rust: Enabled.
                                          // Note: If Interval is 0, Enabled=True has no effect until Interval > 0.
            Interval = 0; // VB6 default. If 0, the timer is disabled regardless of the Enabled property. Rust: Same.
            Left = 0;     // Design-time default. Rust: Same.
            Top = 0;      // Design-time default. Rust: Same.
        }
    }
}