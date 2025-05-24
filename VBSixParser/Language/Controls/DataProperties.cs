// Namespace: VB6Parse.Language.Controls
// For VB6Color, add: using VB6Parse.Language;
// For common enums and data specific enums.

namespace VB6Parse.Language.Controls
{
    /// <summary>
    /// Properties specific to a VB6 Data control.
    /// Provides data binding capabilities to other controls on a form.
    /// </summary>
    public class DataProperties
    {
        public Align Align { get; set; }
        public Appearance Appearance { get; set; }
        public Vb6Color BackColor { get; set; }
        public BOFActionType BOFAction { get; set; }
        public string Caption { get; set; }
        public ConnectType Connect { get; set; } // VB6 Property Name: Connect
        public string DatabaseName { get; set; }
        public CursorTypeDefault DefaultCursorType { get; set; } // VB6 Property Name: DefaultCursorType
        public DataAccessType DefaultType { get; set; } // VB6 Property Name: DefaultType
        // public byte[] DragIcon { get; set; }
        public DragMode DragMode { get; set; }
        public Activation Enabled { get; set; }
        public EOFActionType EOFAction { get; set; }
        public bool Exclusive { get; set; }
        // public Font Font { get; set; } // Separate object
        public Vb6Color ForeColor { get; set; }
        public int Height { get; set; }
        public int Left { get; set; }
        // public byte[] MouseIcon { get; set; }
        public MousePointer MousePointer { get; set; }
        public bool Negotiate { get; set; } // Note: Rust has 'negotitate'
        public OLEDropMode OLEDropMode { get; set; }
        public int Options { get; set; } // Bitmask: dbDenyWrite, dbDenyRead, dbReadOnly, dbAppendOnly etc.
        public bool ReadOnly { get; set; } // Recordset ReadOnly
        public RecordsetTypeEnum RecordsetType { get; set; } // VB6 Property Name: RecordsetType
        public string RecordSource { get; set; }
        public TextDirection RightToLeft { get; set; } // From CommonControlEnums
        public string ToolTipText { get; set; }
        public int Top { get; set; }
        public Visibility Visible { get; set; }
        public int WhatsThisHelpId { get; set; }
        public int Width { get; set; }

        public DataProperties()
        {
            Align = Align.None; // VB6 default. Rust: Same.
            Appearance = Appearance.ThreeD; // VB6 default. Rust: Same.
            BackColor = Vb6Color.FromOleColor(0x8000000F); // System Button Face. Rust: System index 5 (Window Background).
                                                         // Data control usually has button face color.
            BOFAction = BOFActionType.MoveFirst; // VB6 default. Rust: Same.
            Caption = string.Empty; // VB6 Default: "Data1" (or control name). Rust: Empty.
                                    // Let's use control name as default, set by constructor of DataControl.
            Connect = ConnectType.Access; // VB6 Default: "Access". Rust: Same. (If DatabaseName is .mdb)
                                          // More precisely, VB6 default is "" which implies Jet.
            DatabaseName = string.Empty;
            DefaultCursorType = CursorTypeDefault.DefaultCursor; // VB6 default. Rust: Same.
            DefaultType = DataAccessType.UseJet; // VB6 default. Rust: Same (UseJet = 2).
            // DragIcon = null;
            DragMode = DragMode.Manual;
            Enabled = Activation.Enabled;
            EOFAction = EOFActionType.MoveLast; // VB6 default. Rust: Same.
            Exclusive = false; // VB6 default. Rust: Same.
            // Font: Default system font for captions.
            ForeColor = Vb6Color.FromOleColor(0x80000012); // System Button Text. Rust: System index 8 (Window Text).
            Height = 345;  // VB6 default for Data Control. Rust: 1215 (seems very large).
            Left = 0;      // VB6 default. Rust: 480.
            // MouseIcon = null;
            MousePointer = MousePointer.Default;
            Negotiate = false; // VB6 default. Rust: Same.
            OLEDropMode = OLEDropMode.None; // VB6 default. Rust: Same.
            Options = 0; // VB6 default. Rust: Same.
            ReadOnly = false; // VB6 default. Rust: Same.
            RecordsetType = RecordsetTypeEnum.Dynaset; // VB6 default. Rust: Same.
            RecordSource = string.Empty;
            RightToLeft = TextDirection.LeftToRight; // VB6 default. Rust: Same.
            ToolTipText = string.Empty;
            Top = 0;       // VB6 default. Rust: 840.
            Visible = Visibility.Visible;
            WhatsThisHelpId = 0;
            Width = 2055;  // VB6 default for Data Control. Rust: 1140.
        }
    }
}