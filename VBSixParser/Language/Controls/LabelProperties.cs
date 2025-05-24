// Namespace: VB6Parse.Language.Controls
// For VB6Color, if not in the same namespace, add: using VB6Parse.Language;
// For common enums like Alignment, Appearance, etc. they are in this namespace.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Determines if a Label control will wrap text.
    /// </summary>
    public enum WordWrap
    {
        /// <summary>The Label control will not wrap text. (Default)</summary>
        NonWrapping = 0, // VB False
        /// <summary>The Label control will wrap text.</summary>
        Wrapping = -1    // VB True
    }

    /// <summary>
    /// Properties specific to a VB6 Label control.
    /// </summary>
    public class LabelProperties
    {
        public Alignment Alignment { get; set; }
        public Appearance Appearance { get; set; }
        public AutoSize AutoSize { get; set; }
        public Vb6Color BackColor { get; set; }
        public BackStyle BackStyle { get; set; }
        public BorderStyle BorderStyle { get; set; }
        public string Caption { get; set; }
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // Not directly in VB6 Label, but in Rust struct. Check if needed.
                                                // VB6 Label has DataFormat property but it's an object (StdDataFormat).
                                                // For now, string is a placeholder if it's simple text format.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data. Rust uses image::DynamicImage.
                                             // In C#, System.Drawing.Image or byte[] for raw data.
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int Left { get; set; }   // Typically in Twips
        public string LinkItem { get; set; }
        public LinkMode LinkMode { get; set; } // General LinkMode, not FormLinkMode
        public int LinkTimeout { get; set; } // In milliseconds
        public string LinkTopic { get; set; }
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDropMode OLEDropMode { get; set; } // Note: OLEDropMode might not be on all standard Labels unless it's an OLE container.
                                                   // Keeping for parity with Rust struct.
        public TextDirection RightToLeft { get; set; }
        public int TabIndex { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public bool UseMnemonic { get; set; } // If true, '&' in caption creates an access key.
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips
        public WordWrap WordWrap { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="LabelProperties"/> class with default VB6 values.
        /// </summary>
        public LabelProperties()
        {
            Alignment = Alignment.LeftJustify;
            Appearance = Appearance.ThreeD; // VB6 default for Label is 3D
            AutoSize = AutoSize.Fixed;    // Default is False (Fixed) for Label, if True, it resizes to caption.
            BackColor = Vb6Color.FromOleColor(0x8000000F); // System Button Face
            BackStyle = BackStyle.Opaque; // Default for Label
            BorderStyle = BorderStyle.None; // Default for Label
            Caption = string.Empty; // Default is "LabelN" but empty for programmatic.
            DataField = string.Empty;
            // DataFormat = string.Empty; // If we keep it
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000012); // System Button Text
            Height = 255; // Example default, VB6 varies by font. Rust uses 30.
            Left = 0;     // Example default. Rust uses 30.
            LinkItem = string.Empty;
            LinkMode = LinkMode.None;
            LinkTimeout = 5000; // VB6 default is 5000ms (5 seconds). Rust uses 50.
            LinkTopic = string.Empty;
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDropMode = OLEDropMode.None; // Default
            RightToLeft = TextDirection.LeftToRight; // Default is False (LeftToRight)
            TabIndex = 0; // Default for first control, then increments.
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust uses 30.
            UseMnemonic = true; // Default
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1215; // Example default. Rust uses 100.
            WordWrap = WordWrap.NonWrapping; // Default is False (NonWrapping)
        }
    }
}