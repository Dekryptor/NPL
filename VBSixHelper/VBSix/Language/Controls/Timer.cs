using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties specific to a VB6 Timer control.
    /// The Timer control is invisible at runtime and is used to execute code at regular intervals.
    /// </summary>
    public class TimerProperties
    {
        /// <summary>
        /// Determines if the Timer control is active and firing Timer events.
        /// VB6 Property: Enabled. Default is True for a newly added Timer.
        /// </summary>
        public Activation Enabled { get; set; } = Controls.Activation.Enabled; 

        /// <summary>
        /// The interval, in milliseconds, between Timer events.
        /// A value of 0 disables the timer. Maximum value is 65,535 (approx 65.5 seconds).
        /// VB6 Property: Interval. Default is 0.
        /// </summary>
        public int Interval { get; set; } = 0;

        /// <summary>
        /// The Left coordinate of the Timer control icon at design time.
        /// This has no effect at runtime as the Timer is invisible.
        /// VB6 Property: Left.
        /// </summary>
        public int Left { get; set; } = 120; // Typical design-time default Left in Twips

        /// <summary>
        /// The Top coordinate of the Timer control icon at design time.
        /// This has no effect at runtime.
        /// VB6 Property: Top.
        /// </summary>
        public int Top { get; set; } = 120;  // Typical design-time default Top in Twips

        // Standard properties like Name, Index, Tag are part of the parent VB6Control.
        // Timer controls do not have visual properties like BackColor, ForeColor, Font,
        // BorderStyle, Caption, Height, Width (fixed small icon at design time), etc.

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a Timer control.
        /// </summary>
        public TimerProperties() 
        {
            // VB6 IDE often defaults a new Timer's Enabled to True and Interval to 0.
        }

        /// <summary>
        /// Initializes TimerProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public TimerProperties(IDictionary<string, PropertyValue> rawProps)
        {
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            Interval = rawProps.GetInt32("Interval", Interval);
            Left = rawProps.GetInt32("Left", Left);
            // Top is often not explicitly saved in .frm if it's at a default position with Left.
            // If Top is missing from rawProps, it will retain the default C# constructor value.
            Top = rawProps.GetInt32("Top", Top); 
        }
    }
}