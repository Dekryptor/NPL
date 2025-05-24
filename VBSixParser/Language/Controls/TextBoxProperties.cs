// Namespace: VB6Parse.Language.Controls
// For VB6Color, if not in the same namespace, add: using VB6Parse.Language;
// For common enums like Alignment, Appearance, etc. they are in this namespace.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Determines whether a TextBox control has horizontal or vertical scroll bars.
    /// For a TextBox, the MultiLine property must be true for ScrollBars other than None.
    /// Reference: https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa445672(v=vs.60)
    /// </summary>
    public enum ScrollBarsSetting // Renamed from ScrollBars to avoid conflict with System.Windows.Forms.ScrollBars
    {
        /// <summary>No scroll bars are displayed. (Default)</summary>
        None = 0,
        /// <summary>A horizontal scroll bar is displayed.</summary>
        Horizontal = 1,
        /// <summary>A vertical scroll bar is displayed.</summary>
        Vertical = 2,
        /// <summary>Both horizontal and vertical scroll bars are displayed.</summary>
        Both = 3
    }

    /// <summary>
    /// Determines if a TextBox control is multi-line or single-line.
    /// </summary>
    public enum MultiLineSetting // Renamed from MultiLine to avoid conflict if a property is named MultiLine
    {
        /// <summary>The TextBox control is a single-line text box. (Default)</summary>
        SingleLine = 0, // VB False
        /// <summary>The TextBox control is a multi-line text box.</summary>
        MultiLine = -1   // VB True
    }

    /// <summary>
    /// Properties specific to a VB6 TextBox control.
    /// </summary>
    public class TextBoxProperties
    {
        public Alignment Alignment { get; set; }
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public BorderStyle BorderStyle { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public string DataField { get; set; }
        // public string DataFormat { get; set; } // StdDataFormat object in VB6. String for simplicity if basic.
        public string DataMember { get; set; }
        public string DataSource { get; set; }
        // public byte[] DragIcon { get; set; } // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int HelpContextId { get; set; }
        public bool HideSelection { get; set; } // True to hide selection when control loses focus.
        public int Left { get; set; }   // Typically in Twips
        public string LinkItem { get; set; }
        public LinkMode LinkMode { get; set; }
        public int LinkTimeout { get; set; }
        public string LinkTopic { get; set; }
        public bool Locked { get; set; } // True if text cannot be edited.
        public int MaxLength { get; set; } // Maximum number of characters; 0 for no limit.
        // public byte[] MouseIcon { get; set; } // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public MultiLineSetting MultiLine { get; set; }
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public char? PasswordChar { get; set; } // Character to display instead of actual text. Null for none.
        public TextDirection RightToLeft { get; set; }
        public ScrollBarsSetting ScrollBars { get; set; }
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string Text { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="TextBoxProperties"/> class with default VB6 values.
        /// </summary>
        public TextBoxProperties()
        {
            Alignment = Alignment.LeftJustify;
            Appearance = Appearance.ThreeD;
            BackColor = Vb6Color.FromOleColor(0x80000005); // System Window Background
            BorderStyle = BorderStyle.FixedSingle; // Default for TextBox
            CausesValidation = CausesValidation.Yes;
            DataField = string.Empty;
            // DataFormat = string.Empty;
            DataMember = string.Empty;
            DataSource = string.Empty;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000008); // System Window Text
            Height = 285; // Example default, VB6 varies. Rust uses 30.
            HelpContextId = 0;
            HideSelection = true; // Default
            Left = 0;     // Example default. Rust uses 30.
            LinkItem = string.Empty;
            LinkMode = LinkMode.None;
            LinkTimeout = 50; // Rust default. VB6 might be 5000ms.
            LinkTopic = string.Empty;
            Locked = false; // Default
            MaxLength = 0;  // Default (no limit)
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            MultiLine = MultiLineSetting.SingleLine; // Default is False
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // Default for TextBox
            PasswordChar = null; // Default (no password char)
            RightToLeft = TextDirection.LeftToRight; // Default is False
            ScrollBars = ScrollBarsSetting.None; // Default
            TabIndex = 0;
            TabStop = TabStop.Included; // Default is True
            Text = string.Empty; // Default is "TextN" but empty for programmatic.
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust uses 30.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1215; // Example default. Rust uses 100.
        }
    }
}