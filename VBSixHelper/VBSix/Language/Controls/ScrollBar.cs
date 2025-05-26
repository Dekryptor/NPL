using VBSix.Language; // For VB6Color (though ScrollBar doesn't use Back/ForeColor directly)
using VBSix.Utilities; // For PropertyParserHelpers extension methods
using VBSix.Parsers;   // For PropertyValue
using System.Collections.Generic; // For IDictionary

namespace VBSix.Language.Controls
{
    /// <summary>
    /// Represents the properties common to VB6 HScrollBar (Horizontal ScrollBar)
    /// and VScrollBar (Vertical ScrollBar) controls.
    /// </summary>
    public class ScrollBarProperties
    {
        public CausesValidation CausesValidation { get; set; } = Controls.CausesValidation.Yes;
        // DragIcon, MouseIcon are FRX-based.
        // Their resolved byte[] data will be stored in the VB6HScrollBarKind/VB6VScrollBarKind classes.
        public DragMode DragMode { get; set; } = Controls.DragMode.Manual;
        public Activation Enabled { get; set; } = Controls.Activation.Enabled;
        public int Height { get; set; } = 255;  // Default height for HScrollBar in Twips; VScrollBar width is similar
        public int HelpContextID { get; set; } = 0;
        public int LargeChange { get; set; } = 1;    // Amount Value changes when clicking between thumb and arrow
        public int Left { get; set; } = 120;   // Typical default Left in Twips
        public int Max { get; set; } = 32767;  // Maximum scroll value
        public int Min { get; set; } = 0;      // Minimum scroll value
        public MousePointer MousePointer { get; set; } = Controls.MousePointer.Default;
        public TextDirection RightToLeft { get; set; } = Controls.TextDirection.LeftToRight; // Affects HScrollBar appearance
        public int SmallChange { get; set; } = 1;    // Amount Value changes when clicking an arrow
        public int TabIndex { get; set; } = 0; // Placeholder
        public TabStop TabStop { get; set; } = Controls.TabStop.Included; // Scrollbars can receive focus
        public int Top { get; set; } = 120;    // Typical default Top in Twips
        public int Value { get; set; } = 0;      // Current position of the scroll box (thumb)
        public Visibility Visible { get; set; } = Controls.Visibility.Visible;
        public int WhatsThisHelpID { get; set; } = 0;
        public int Width { get; set; } = 2055;   // Default width for HScrollBar in Twips; VScrollBar height is similar

        // Note: ScrollBars do not have BackColor, ForeColor, or Font properties directly settable in the IDE.
        // Their appearance is usually system-defined.

        /// <summary>
        /// Default constructor initializing with common VB6 defaults for a ScrollBar control.
        /// Specific defaults for Height/Width might differ slightly between HScrollBar and VScrollBar.
        /// </summary>
        public ScrollBarProperties() { }

        /// <summary>
        /// Initializes ScrollBarProperties by parsing values from a raw property dictionary.
        /// </summary>
        /// <param name="rawProps">The dictionary of raw property names and values.</param>
        public ScrollBarProperties(IDictionary<string, PropertyValue> rawProps)
        {
            CausesValidation = rawProps.GetEnum("CausesValidation", CausesValidation);
            DragMode = rawProps.GetEnum("DragMode", DragMode);
            Enabled = rawProps.GetEnum("Enabled", Enabled);
            Height = rawProps.GetInt32("Height", Height);
            HelpContextID = rawProps.GetInt32("HelpContextID", HelpContextID);
            LargeChange = rawProps.GetInt32("LargeChange", LargeChange);
            Left = rawProps.GetInt32("Left", Left);
            Max = rawProps.GetInt32("Max", Max);
            Min = rawProps.GetInt32("Min", Min);
            MousePointer = rawProps.GetEnum("MousePointer", MousePointer);
            RightToLeft = rawProps.GetEnum("RightToLeft", RightToLeft);
            SmallChange = rawProps.GetInt32("SmallChange", SmallChange);
            TabIndex = rawProps.GetInt32("TabIndex", TabIndex);
            TabStop = rawProps.GetEnum("TabStop", TabStop);
            Top = rawProps.GetInt32("Top", Top);
            Value = rawProps.GetInt32("Value", Value);
            Visible = rawProps.GetEnum("Visible", Visible);
            WhatsThisHelpID = rawProps.GetInt32("WhatsThisHelpID", WhatsThisHelpID);
            Width = rawProps.GetInt32("Width", Width);
        }
    }
}