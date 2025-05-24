// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Frame control.
    /// A Frame is a container used to group other controls and has a visible border and caption.
    /// </summary>
    public class FrameProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; } // Background color of the frame's client area.
        public BorderStyle BorderStyle { get; set; } // Style of the frame's border (None or FixedSingle).
        public string Caption { get; set; } // Text displayed in the upper border of the frame.
        public ClipControls ClipControls { get; set; } // Determines if drawing/controls are clipped to the frame.
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; } // Color of the caption text.
        // Note: Font object properties (Name, Size, Bold, etc.) for the caption are also part of a Frame.
        // These would typically be represented by a nested FontProperties class.
        public int Height { get; set; } // Outer height in Twips.
        public int HelpContextId { get; set; }
        public int Left { get; set; }   // Outer left edge in Twips, relative to container.
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDropMode OLEDropMode { get; set; } // Whether the control can act as an OLE drop target.
        public TextDirection RightToLeft { get; set; } // For caption text direction and control layout.
        public int TabIndex { get; set; } // Position in tab order.
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Outer top edge in Twips, relative to container.
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Outer width in Twips.

        /// <summary>
        /// Initializes a new instance of the <see cref="FrameProperties"/> class with default VB6 values.
        /// </summary>
        public FrameProperties()
        {
            Appearance = Appearance.ThreeD; // Default VB6 appearance for Frame.
            BackColor = Vb6Color.FromOleColor(0x8000000F); // vbButtonFace (system color for control backgrounds).
            BorderStyle = BorderStyle.FixedSingle; // Default.
            Caption = "Frame1"; // Default caption.
            ClipControls = ClipControls.Clipped; // Default for Frame (VB6 True, Rust 0).
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000012); // vbButtonText (system color for text on controls).
            // Font typically defaults to system font like MS Sans Serif, 8pt.
            Height = 1500; // Example default height (Rust's 30 is too small).
            HelpContextId = 0;
            Left = 120;    // Example default.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDropMode = OLEDropMode.None; // Default.
            RightToLeft = TextDirection.LeftToRight; // Default is False.
            TabIndex = 0; // Set by IDE based on creation order.
            ToolTipText = string.Empty;
            Top = 120;     // Example default.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 3000;  // Example default width (Rust's 100 is too small).
        }
    }
}