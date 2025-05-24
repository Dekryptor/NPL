// Namespace: VB6Parse.Language.Controls
using System;
using System.Collections.Generic;
using System.Linq;
using VB6Parse.Model;
using VB6Parse.Utilities;

namespace VB6Parse.Language.Controls
{
    public class TimerProperties : ControlSpecificPropertiesBase
    {
        public Activation Enabled { get; set; } = Activation.Enabled; // Or False by default often
        public int Index { get; set; } = -1; // Contextual, for timer control arrays
        public int Interval { get; set; } = 0; // In milliseconds. 0 disables the timer.
        public int Left { get; set; } = 0; // Position of the icon on the form at design time
        public string Tag { get; set; } = string.Empty;
        public int Top { get; set; } = 0;  // Position of the icon on the form at design time

        // Timer controls are non-visual and do not have properties like
        // BackColor, ForeColor, Caption, Font, Height, Width, TabIndex, TabStop, Visible, etc.
        // The base Font property will be nullified.

        public TimerProperties()
        {
            // Timers don't have their own Font property in a meaningful way for runtime.
            // Setting Font to null in the base or overriding PopulateFontProperty.
            this.Font = null;
        }

        public override void PopulateFromDictionaries(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyDictionary<string, byte[]> binaryProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            Enabled = PropertyParsingHelpers.GetEnum(textualProps, "Enabled", this.Enabled);
            // Index is typically handled by the Vb6Control's main properties, but if "Index"
            // specifically appears in textualProps for a Timer, it could be captured here too.
            // However, ControlParseState and BuildControlFromState already handle Index based on the name.
            // If Index is explicitly set as a property like `Index = 0`, this would catch it.
            if (textualProps.TryGetValue("Index", out string indexValStr) && int.TryParse(indexValStr, out int parsedIndex))
            {
                 // This would only be for an explicit "Index = X" line, not the (0) in Name(0).
                 // The Vb6Control.Index property is the primary one.
            }

            Interval = PropertyParsingHelpers.GetInt32(textualProps, "Interval", this.Interval);
            Left = PropertyParsingHelpers.GetInt32(textualProps, "Left", this.Left);
            Tag = PropertyParsingHelpers.GetString(textualProps, "Tag", this.Tag);
            Top = PropertyParsingHelpers.GetInt32(textualProps, "Top", this.Top);

            // No binary properties expected for a standard Timer.
            // No visual properties like Font, BackColor, ForeColor, Height, Width.
        }

        protected override void PopulateFontProperty(
            IReadOnlyDictionary<string, string> textualProps,
            IReadOnlyList<Vb6PropertyGroup> propertyGroups)
        {
            // Timers do not have a Font property.
            this.Font = null;
        }
    }
}