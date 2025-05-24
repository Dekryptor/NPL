// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 PictureBox control.
    /// PictureBoxes can display images, act as containers, and serve as drawing surfaces.
    /// </summary>
    public class PictureBoxProperties
    {
        public Align Align { get; set; }
        public Appearance Appearance { get; set; }
        public AutoRedraw AutoRedraw { get; set; } // If True, graphics are persistent.
        public AutoSize AutoSize { get; set; } // If True, resizes to fit its Picture.
        public Vb6Color BackColor { get; set; }
        public BorderStyle BorderStyle { get; set; } // None or FixedSingle.
        public CausesValidation CausesValidation { get; set; } // Shared with other controls.
        public ClipControls ClipControls { get; set; } // If True, graphics/controls are clipped.
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public DrawMode DrawMode { get; set; } // For graphics methods.
        public DrawStyle DrawStyle { get; set; } // For graphics methods.
        public int DrawWidth { get; set; } // Line thickness for graphics methods.
        public Activation Enabled { get; set; }
        public Vb6Color FillColor { get; set; } // For graphics methods.
        public FillStyleType FillStyle { get; set; } // For graphics methods (uses FillStyleType enum).
        public FontTransparency FontTransparent { get; set; } // If True, text background is transparent.
        // Note: Font object properties (Name, Size, Bold, etc.) are also part of a PictureBox.
        public Vb6Color ForeColor { get; set; } // For text and graphics methods.
        public HasDeviceContext HasDC { get; set; } // Always True for PictureBox.
        public int Height { get; set; } // Outer height in Twips.
        public int HelpContextId { get; set; }
        // public Image Image { get; set; } // Current image displayed (runtime only, not saved in .frm like Picture)
        public int Left { get; set; }   // Outer left edge in Twips.
        public string LinkItem { get; set; } // DDE item.
        public LinkMode LinkMode { get; set; } // DDE link type.
        public int LinkTimeout { get; set; } // DDE timeout in 1/10 seconds.
        public string LinkTopic { get; set; } // DDE topic.
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public bool NegotiatePosition { get; set; } // VB6 Negotiate property (Rust: negotiate)
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public byte[] Picture { get; set; } // Image data loaded at design time (from .frx).
        public TextDirection RightToLeft { get; set; }
        public float ScaleHeight { get; set; } // User-defined height for custom coordinate system.
        public float ScaleLeft { get; set; }   // User-defined left for custom coordinate system.
        public ScaleMode ScaleMode { get; set; } // Unit for custom coordinate system.
        public float ScaleTop { get; set; }    // User-defined top for custom coordinate system.
        public float ScaleWidth { get; set; }  // User-defined width for custom coordinate system.
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; } // Typically True for PictureBox.
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Outer top edge in Twips.
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Outer width in Twips.

        public PictureBoxProperties()
        {
            Align = Align.None;
            Appearance = Appearance.ThreeD;
            AutoRedraw = AutoRedraw.Manual; // VB6 default is False (0). Rust: Manual (0).
            AutoSize = AutoSize.Fixed;     // VB6 default is False (0). Rust: Fixed (0).
            BackColor = Vb6Color.FromOleColor(0x8000000F); // vbButtonFace. Rust: Same.
            BorderStyle = BorderStyle.FixedSingle; // VB6 default. Rust: Same.
            CausesValidation = CausesValidation.Yes;
            ClipControls = ClipControls.Clipped; // VB6 default is True (-1). Rust: Clipped (0).
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            DrawMode = DrawMode.CopyPen; // VB6 default. Rust: Same.
            DrawStyle = DrawStyle.Solid; // VB6 default. Rust: Same.
            DrawWidth = 1; // VB6 default. Rust: Same.
            Enabled = Activation.Enabled;
            FillColor = Vb6Color.FromRGB(0, 0, 0); // vbBlack. Rust: Same.
            FillStyle = FillStyleType.Transparent; // VB6 default. Rust: Transparent (1).
            FontTransparent = FontTransparency.Transparent; // VB6 default is True (-1). Rust: Transparent (1).
            ForeColor = Vb6Color.FromOleColor(0x80000012); // vbButtonText. Rust: Same.
            HasDC = HasDeviceContext.Yes; // Always True for PictureBox. Rust: Yes (1).
            Height = 1200; // Example default (Rust's 30 is too small).
            HelpContextId = 0;
            // Image = null;
            Left = 120;    // Example default.
            LinkItem = string.Empty;
            LinkMode = LinkMode.None; // VB6 default. Rust: Same.
            LinkTimeout = 50; // VB6 default (5 seconds). Rust: Same.
            LinkTopic = string.Empty;
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            NegotiatePosition = false; // VB6 default. Rust: Same.
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: None (0).
            Picture = null; // No picture by default.
            RightToLeft = TextDirection.LeftToRight;
            // Scale properties default to the control's dimensions in Twips if ScaleMode is Twip.
            ScaleHeight = 0; // Actual value set by VB at runtime based on Height and ScaleMode
            ScaleLeft = 0;
            ScaleMode = ScaleMode.Twip; // VB6 default. Rust: Same.
            ScaleTop = 0;
            ScaleWidth = 0;  // Actual value set by VB at runtime based on Width and ScaleMode
            TabIndex = 0;
            TabStop = TabStop.Included; // VB6 default is True. Rust: Included (0).
            ToolTipText = string.Empty;
            Top = 120;     // Example default.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1800;  // Example default (Rust's 100 is too small).
        }
    }
}