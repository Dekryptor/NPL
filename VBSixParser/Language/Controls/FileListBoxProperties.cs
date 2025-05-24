// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 FileListBox control.
    /// Displays a list of files from a specified path, matching a pattern and attributes.
    /// </summary>
    public class FileListBoxProperties
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
        public MultiSelect MultiSelect { get; set; } // How many items can be selected.
        public OLEDragMode OLEDragMode { get; set; }
        public OLEDropMode OLEDropMode { get; set; }
        public int TabIndex { get; set; }
        public TabStop TabStop { get; set; }
        public string ToolTipText { get; set; }
        public int Top { get; set; }
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }

        // FileListBox-specific properties
        public string Path { get; set; } // Current path being displayed. Write-only at design time.
        public string FileName { get; set; } // Selected file name. Can be set to select a file.
        public string Pattern { get; set; } // File pattern (e.g., "*.txt").

        public bool Archive { get; set; } // True to display files with Archive attribute.
        public bool Hidden { get; set; }  // True to display files with Hidden attribute.
        public bool Normal { get; set; }  // True to display files with no other attributes set.
        public bool ReadOnly { get; set; } // True to display files with ReadOnly attribute.
        public bool System { get; set; }   // True to display files with System attribute.

        // Runtime properties not typically in .frm but important for behavior:
        // ListCount, List(index), Selected(index), ListIndex

        public FileListBoxProperties()
        {
            Appearance = Appearance.ThreeD; // VB6 default
            BackColor = Vb6Color.FromOleColor(0x80000005); // vbWindowBackground. Rust: System index 5.
            CausesValidation = CausesValidation.Yes;
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            ForeColor = Vb6Color.FromOleColor(0x80000008); // vbWindowText. Rust: System index 8.
            // Font: MS Sans Serif, 8pt typically.
            Height = 1935; // VB6 default Height in Twips (approx 12 lines + border)
                           // Rust default is 1260
            HelpContextId = 0;
            Left = 0;      // VB6 default. Rust: 710
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            MultiSelect = MultiSelect.None; // VB6 default. Rust: Same.
            OLEDragMode = OLEDragMode.Manual;
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: Same.
            TabIndex = 0; // Set by IDE.
            TabStop = TabStop.Included; // VB6 default is True. Rust: Included (0).
            ToolTipText = string.Empty;
            Top = 0;       // VB6 default. Rust: 480
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 1875;  // VB6 default Width in Twips. Rust: 735

            // FileListBox specific defaults:
            Path = string.Empty; // VB6 initializes Path to current directory at runtime if empty.
                                 // Design-time Path property is write-only.
            FileName = string.Empty; // No file selected initially.
            Pattern = "*.*";     // VB6 default. Rust: Same.

            Archive = true;      // VB6 default. Rust: Include (-1).
            Hidden = false;      // VB6 default. Rust: Exclude (0).
            Normal = true;       // VB6 default. Rust: Include (-1).
            ReadOnly = false;    // VB6 default (shows read-only files by default if ReadOnly property is false).
                                 // Let's re-check VB6 behavior for ReadOnly attribute.
                                 // VB6 FileListBox.ReadOnly property: "Determines if read-only files are displayed."
                                 // Default is False, meaning read-only files *are* displayed.
                                 // If True, only read-only files are displayed. This is confusing.
                                 // The Rust `ReadOnlyAttribute::Include` (-1) default for `read_only` field matches VB6's `File1.ReadOnly = True`
                                 // which means "display files that have the read-only attribute".
                                 // However, the VB6 property `FileListBox.ReadOnly` (which sounds like it should correspond to this)
                                 // actually behaves differently: if `FileListBox.ReadOnly = True`, it *only* lists read-only files.
                                 // If `FileListBox.ReadOnly = False` (default), it lists files regardless of their read-only status (matching `Normal=True`).
                                 // The Rust model seems to map to individual boolean flags for each attribute: ShowArchive, ShowHidden etc.
                                 // Let's stick to the VB6 property names for these booleans.
                                 // VB6 FileListBox Properties: Archive, Hidden, Normal, ReadOnly, System (all boolean)
                                 // Default values in VB6 IDE:
                                 // Archive = True
                                 // Hidden = False
                                 // Normal = True
                                 // ReadOnly = True  (This means "include files that are read-only" in the list)
                                 // System = False
            System = false;      // VB6 default. Rust `SystemAttribute::Include` for `system` field is different.
                                 // Let's use VB6 defaults for these boolean properties.
                                 // Re-aligning with VB6 IDE defaults:
            ReadOnly = true;     // VB6 default. Rust: Include (-1).
            System = false;      // VB6 default. Rust `system` field defaults to `Include` (-1).
        }
    }
}