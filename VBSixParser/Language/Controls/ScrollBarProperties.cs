// Namespace: VB6Parse.Language.Controls
// For common enums like Activation, DragMode, etc.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to VB6 HScrollBar and VScrollBar controls.
    /// </summary>
    public class ScrollBarProperties
    {
        public CausesValidation CausesValidation { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int HelpContextId { get; set; }
        public int LargeChange { get; set; } // Amount Value changes when clicking in the scroll bar area
        public int Left { get; set; }   // Typically in Twips
        public int Max { get; set; }    // Maximum scroll value
        public int Min { get; set; }    // Minimum scroll value
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public TextDirection RightToLeft { get; set; } // Affects visual layout (Min/Max positions)
        public int SmallChange { get; set; } // Amount Value changes when clicking an arrow
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public int Value { get; set; }   // Current scroll position
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrollBarProperties"/> class with default VB6 values.
        /// </summary>
        public ScrollBarProperties()
        {
            CausesValidation = CausesValidation.Yes;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            // Height/Width defaults depend on whether it's HScrollBar or VScrollBar
            // HScrollBar: Height is fixed-ish (e.g., 255-300 Twips), Width varies.
            // VScrollBar: Width is fixed-ish (e.g., 255-300 Twips), Height varies.
            // Rust defaults (Height=30, Width=100) are very small.
            // Let's use more typical HScrollBar defaults and adjust in specific control constructors if needed.
            Height = 285; // Typical height for an HScrollBar
            Width = 1200; // Example width for an HScrollBar
            HelpContextId = 0;
            LargeChange = 1; // Default
            Left = 0;      // Example default (Rust: 30)
            Max = 32767;   // Default
            Min = 0;       // Default
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            RightToLeft = TextDirection.LeftToRight; // Default
            SmallChange = 1; // Default
            TabIndex = 0;    // Typically assigned by IDE
            TabStop = TabStop.Included; // Default is True
            Top = 0;       // Example default (Rust: 30)
            Value = 0;       // Default
            Visible = Visibility.Visible; // Default
            WhatsThisHelpId = 0;
        }
    }
}