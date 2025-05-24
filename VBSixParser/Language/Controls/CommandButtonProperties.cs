// Namespace: VB6Parse.Language.Controls
// For VB6Color, if not in the same namespace, add: using VB6Parse.Language;
// For common enums like Appearance, DragMode, etc. they are in this namespace.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 CommandButton control.
    /// </summary>
    public class CommandButtonProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public bool Cancel { get; set; } // True if this button is activated when Esc is pressed.
        public string Caption { get; set; }
        public CausesValidation CausesValidation { get; set; }
        public bool Default { get; set; } // True if this button is activated when Enter is pressed.
        // public byte[] DisabledPicture { get; set; } // Placeholder for image data
        // public byte[] DownPicture { get; set; }     // Placeholder for image data
        // public byte[] DragIcon { get; set; }        // Placeholder for image data
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public int Height { get; set; } // Typically in Twips
        public int HelpContextId { get; set; }
        public int Left { get; set; }   // Typically in Twips
        public Vb6Color MaskColor { get; set; } // Used with Picture property if Style is Graphical
        // public byte[] MouseIcon { get; set; }       // Placeholder for image data
        public MousePointer MousePointer { get; set; }
        public OLEDropMode OLEDropMode { get; set; } // For OLE drag/drop operations
        // public byte[] Picture { get; set; }         // Placeholder for image data
        public TextDirection RightToLeft { get; set; } // For BiDi text display
        public Style Style { get; set; } // Standard or Graphical
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }     // Typically in Twips
        public UseMaskColor UseMaskColor { get; set; } // If MaskColor should be used for transparency with Picture
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }   // Typically in Twips

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandButtonProperties"/> class with default VB6 values.
        /// </summary>
        public CommandButtonProperties()
        {
            Appearance = Appearance.ThreeD;
            BackColor = Vb6Color.FromOleColor(0x8000000F); // System Button Face
            Cancel = false;
            Caption = string.Empty; // Default is "CommandN" but empty for programmatic.
            CausesValidation = CausesValidation.Yes; // Default
            Default = false;
            // DisabledPicture = null;
            // DownPicture = null;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            Height = 375; // Example default, VB6 varies. Rust uses 30.
            HelpContextId = 0;
            Left = 0;     // Example default. Rust uses 30.
            MaskColor = Vb6Color.FromOleColor(0x00C0C0C0); // Light Gray (Silver), Rust uses this.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDropMode = OLEDropMode.None; // Default
            // Picture = null;
            RightToLeft = TextDirection.LeftToRight; // Default is False
            Style = Style.Standard; // Default
            TabIndex = 0; // Default for first control
            TabStop = TabStop.Included; // Default is True
            ToolTipText = string.Empty;
            Top = 0;      // Example default. Rust uses 30.
            UseMaskColor = UseMaskColor.DoNotUseMaskColor; // Default is False
            WhatsThisHelpId = 0;
            Width = 1215; // Example default. Rust uses 100.
        }
    }
}