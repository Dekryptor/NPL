// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums and OLE specific enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 OLE Container control.
    /// Used to embed or link OLE objects.
    /// </summary>
    public class OleProperties // VB6 Control Name: OLE
    {
        public Appearance Appearance { get; set; }
        public OLEAutoActivateEnum AutoActivate { get; set; }
        public bool AutoVerbMenu { get; set; }
        public Vb6Color BackColor { get; set; }
        public BackStyle BackStyle { get; set; }
        public BorderStyle BorderStyle { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public string Class { get; set; } // OLE Class (e.g., "Word.Document.8")
        public string DataField { get; set; } // For data binding
        public string DataSource { get; set; } // Name of Data control for binding
        public OLEDisplayTypeEnum DisplayType { get; set; }
        // public byte[] DragIcon { get; set; }
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        // public Font Font { get; set; } // OLE control itself doesn't have Font, but hosted object might.
        public int Height { get; set; }
        public int HelpContextId { get; set; }
        public string HostName { get; set; } // Application name using the OLE object
        public int Left { get; set; }
        public int MiscFlags { get; set; } // OLE miscellaneous flags
        // public byte[] MouseIcon { get; set; }
        public MousePointer MousePointer { get; set; }
        public bool OLEDropAllowed { get; set; }
        public OLETypeAllowedEnum OLETypeAllowed { get; set; }
        public SizeMode SizeMode { get; set; }
        public string SourceDoc { get; set; } // Path to source document file
        public string SourceItem { get; set; } // Item within source document (e.g., cell range)
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public int Top { get; set; }
        public OLEUpdateOptionsEnum UpdateOptions { get; set; }
        public int Verb { get; set; } // Default verb for AutoActivate
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }

        // Other important runtime properties:
        // Object (reference to the OLE object itself)
        // ObjectVerbs, ObjectVerbsCount
        // Action (to create/delete objects, etc.)

        public OleProperties()
        {
            Appearance = Appearance.ThreeD; // VB6 default. Rust: Same.
            AutoActivate = OLEAutoActivateEnum.DoubleClick; // VB6 default. Rust: Same.
            AutoVerbMenu = true; // VB6 default. Rust: Same.
            BackColor = Vb6Color.FromOleColor(0x80000005); // vbWindowBackground. Rust: Same.
            BackStyle = BackStyle.Opaque; // VB6 default. Rust: Same.
            BorderStyle = BorderStyle.FixedSingle; // VB6 default. Rust: Same.
            CausesValidation = CausesValidation.Yes;
            Class = string.Empty; // No class by default.
            DataField = string.Empty;
            DataSource = string.Empty;
            DisplayType = OLEDisplayTypeEnum.Content; // VB6 default. Rust: Same.
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            Height = 3000; // VB6 default (e.g. 250x200 pixels -> 3750x3000 twips). Rust: 375 (very small).
            HelpContextId = 0;
            HostName = string.Empty; // VB6 fills this at runtime.
            Left = 0;      // VB6 default. Rust: 600.
            MiscFlags = 0; // VB6 default. Rust: Same.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDropAllowed = false; // VB6 default. Rust: Same.
            OLETypeAllowed = OLETypeAllowedEnum.Either; // VB6 default. Rust: Same.
            SizeMode = SizeMode.Clip; // VB6 default (vbOLESizeClip). Rust: Same.
            SourceDoc = string.Empty;
            SourceItem = string.Empty;
            TabIndex = 0; // Set by IDE.
            TabStop = TabStop.Included; // VB6 default True. Rust: Same.
            Top = 0;       // VB6 default. Rust: 1200.
            UpdateOptions = OLEUpdateOptionsEnum.Automatic; // VB6 default. Rust: Same.
            Verb = 0;      // Primary verb. VB6 default. Rust: Same.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 3750;  // VB6 default. Rust: 1335.
        }
    }
}