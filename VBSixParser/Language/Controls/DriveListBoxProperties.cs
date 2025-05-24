// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 DriveListBox control.
    /// Displays a list of available disk drives.
    /// </summary>
    public class DriveListBoxProperties
    {
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public CausesValidation CausesValidation { get; set; }
        // public byte[] DragIcon { get; set; } // Not typically used for DriveListBox
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        // Font properties (FontName, FontSize, etc.) are also relevant.
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; } // Height of the control (dropdown list height is automatic)
        public int HelpContextId { get; set; }
        public int Left { get; set; }
        // public byte[] MouseIcon { get; set; } // Not typically used
        public MousePointer MousePointer { get; set; }
        public OLEDropMode OLEDropMode { get; set; } // Usually None
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }

        // DriveListBox-specific property
        /// <summary>
        /// Gets or sets the currently selected drive.
        /// Represented as a string, e.g., "C:", "A:\", or "C:\Windows".
        /// When set, the list updates and selects the specified drive.
        /// </summary>
        public string Drive { get; set; }

        // Runtime properties not typically in .frm but important for behavior:
        // ListCount, List(index), ListIndex (relative to displayed drives)

        public DriveListBoxProperties()
        {
            Appearance = Appearance.ThreeD; // VB6 default
            BackColor = Vb6Color.FromOleColor(0x80000005); // vbWindowBackground. Rust: System index 5.
            CausesValidation = CausesValidation.Yes; // VB6 default
            // DragIcon = null;
            DragMode = DragMode.Manual; // VB6 default
            Enabled = Activation.Enabled; // VB6 default
            ForeColor = Vb6Color.FromOleColor(0x80000008); // vbWindowText. Rust: System index 8.
            // Font: MS Sans Serif, 8pt typically.
            Height = 315;  // VB6 default Height in Twips (standard combo box height). Rust: 319.
            HelpContextId = 0;
            Left = 0;      // VB6 default. Rust: 480.
            // MouseIcon = null;
            MousePointer = MousePointer.Default; // VB6 default
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: Same.
            TabIndex = 0; // Set by IDE.
            TabStop = TabStop.Included; // VB6 default is True. Rust: Included (0).
            ToolTipText = string.Empty;
            Top = 0;       // VB6 default. Rust: 960.
            Visible = Visibility.Visible; // VB6 default
            WhatsThisHelpId = 0;
            Width = 1875;  // VB6 default Width in Twips. Rust: 1455.

            // DriveListBox specific default:
            Drive = string.Empty; // At runtime, VB6 sets this to the current drive.
                                  // Design-time, it's usually empty or reflects dev environment current drive.
        }
    }
}