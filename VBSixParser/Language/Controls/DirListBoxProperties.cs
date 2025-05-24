// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 DirListBox control.
    /// Displays a hierarchical list of directories.
    /// </summary>
    public class DirListBoxProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public CausesValidation CausesValidation { get; set; }
        // public byte[] DragIcon { get; set; }
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        // Font properties (FontName, FontSize, etc.) are also relevant.
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; }
        public int HelpContextId { get; set; }
        public int Left { get; set; }
        // public byte[] MouseIcon { get; set; }
        public MousePointer MousePointer { get; set; }
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }

        // DirListBox-specific property
        /// <summary>
        /// Gets or sets the current path displayed in the DirListBox.
        /// When set, the list updates to show the subdirectories of the new path.
        /// </summary>
        public string Path { get; set; }

        // Runtime properties not typically in .frm but important for behavior:
        // ListCount, List(index), ListIndex

        public DirListBoxProperties()
        {
            Appearance = Appearance.ThreeD; // VB6 default
            BackColor = Vb6Color.FromOleColor(0x80000005); // vbWindowBackground. Rust: System index 5.
            CausesValidation = CausesValidation.Yes;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000008); // vbWindowText. Rust: System index 8.
            // Font: MS Sans Serif, 8pt typically.
            Height = 1935; // VB6 default Height in Twips (approx 12 lines + border). Rust: 3195.
            HelpContextId = 0;
            Left = 0;      // VB6 default. Rust: 720.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: Same.
            TabIndex = 0; // Set by IDE.
            TabStop = TabStop.Included; // VB6 default is True. Rust: Included (0).
            ToolTipText = string.Empty;
            Top = 0;       // VB6 default. Rust: 720.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1875;  // VB6 default Width in Twips. Rust: 975.

            // DirListBox specific default:
            Path = string.Empty; // VB6 initializes Path to current directory at runtime.
                                 // Design-time Path property is effectively write-only or shows current dev path.
        }
    }
}